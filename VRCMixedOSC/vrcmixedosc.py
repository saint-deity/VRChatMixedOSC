port_output = 9000
port_input = 9001
address = "http://127.0.0.1"

import asyncio
import queue, threading, datetime, os, time, textwrap
import speech_recognition as sr

from datetime import timedelta
import time, os

import traceback
import subprocess

from speech_recognition import UnknownValueError, WaitTimeoutError, AudioData
from pythonosc import udp_client
from pythonosc.dispatcher import Dispatcher
from pythonosc.osc_server import BlockingOSCUDPServer
client = udp_client.SimpleUDPClient("127.0.0.1", 9000)

from yaml import load
try:
    from yaml import CLoader as Loader
except ImportError:
    from yaml import Loader
from yaml import dump

from winsdk.windows.media.control import \
    GlobalSystemMediaTransportControlsSessionManager as MediaManager
from winsdk.windows.media.control import \
    GlobalSystemMediaTransportControlsSessionPlaybackStatus
from winsdk.windows.media.core import \
    AudioStreamDescriptor

config_music = {
    'Enabled': True,
    'DisplayFormat': "( NP: {song_artist} - {song_title}{song_position} )",
    'PausedFormat': "Playback Paused",
    'OnlyShowOnChange': True,
    'UseTextFile': False,
    'TextFileLocation': "",
    'TextFileUpdateAlways': False,
    'AFK': False,
    'AFKSince': 0,
}

config_subs = {
    'FollowMicMute': True, 
    'CapturedLanguage': "en-US", 
    'TranscriptionMethod': "Google", 
    'TranscriptionRateLimit': 1200,
    'EnableTranslation': False, 
    'TranslateMethod': "Google", 
    'TranslateToken': "", 
    'TranslateTo': "en-US", 
    'AllowOSCControl': True, 
    'Pause': False, 
    'TranslateInterumResults': True, 
    'OSCControlPort': 9001
}

config_activity = {
    'SubText': "",
}

def load_config():
    global config
    with open("config.yaml", 'r') as stream:
        config = load(stream, Loader=Loader)
    print("[VRCMixedOSC] Loaded config file!")

state = {'selfMuted': False}
state_lock = threading.Lock()

r = sr.Recognizer()
audio_queue = queue.Queue()

last_displayed_song = ("","")
displayed_timestamp = None
last_reported_timestamp = None

textfile_first_tick = False
activity = ''

'''
State Management
'''

def get_state(key):
    global state, state_lock
    state_lock.acquire()
    result = state.copy()
    if key in state:
        result = state[key]
    state_lock.release()
    return result

def set_state(key, value):
    global state, state_lock
    state_lock.acquire()
    state[key] = value
    state_lock.release()

'''
Sound Processing
'''

def process_audio():
    global r, audio_queue, config_subs, client
    current_text = ""
    last_text = ""
    last_time = datetime.datetime.now()

    while True:
        ad, final = audio_queue.get()
        if config_subs['FollowMicMute'] and get_state("selfMuted"):
            continue

        if config_subs['Pause']:
            continue

        # send chatbox to typing
        client.send_message("/chatbox/typing", (not final))
        text = None

        time_now = datetime.datetime.now()
        buffer = time_now - last_time
        if buffer.total_seconds() < 1 and not final:
            continue

        try:
            text = r.recognize_google(ad, language=config_subs['CapturedLanguage'])
        except UnknownValueError:
            #print("Could not understand audio")
            client.send_message("/chatbox/typing", False)
            continue
        except WaitTimeoutError:
            #print("Timeout waiting for audio")
            client.send_message("/chatbox/typing", False)
            continue
        except Exception as e:
            #print("Error processing audio: ", e)
            client.send_message("/chatbox/typing", False)
            continue

        if text is None or text == "":
            continue

        current_text = text
        if current_text == last_text:
            continue

        last_text = current_text
        
        last_time_millis = buffer.total_seconds() * 1000
        if last_time_millis < config_subs['TranscriptionRateLimit']:
            sleep_time = config_subs['TranscriptionRateLimit'] - last_time_millis
            print("Sending too fast, sleeping for ", sleep_time)
            time.sleep(sleep_time / 1000.0)
        
        print("Recognized: ", current_text)
        client.send_message("/chatbox/input", [current_text, True])

#collect_audio
def audio_thread():
    global audio_queue, r, config
    mic = sr.Microphone()
    print("Starting audio thread")
    did = mic.get_pyaudio().PyAudio().get_default_input_device_info()
    print("Using device: ", did.get('name'))
    with mic as source:
        audio_buf = None
        buf_size = 0
        while True:
            audio = None
            try:
                audio = r.listen(source, phrase_time_limit=1, timeout=0.1)
            except WaitTimeoutError:
                if audio_buf is not None:
                    audio_queue.put((audio_buf, True))
                    audio_buf = None
                    buf_size = 0
                continue

            if audio is not None:
                if audio_buf is None:
                    audio_buf = audio
                else:
                    buf_size += 1
                    if buf_size > 10:
                        audio_buf = audio
                        buf_size = 0
                    else:
                        audio_buf = AudioData(audio_buf.frame_data + audio.frame_data, audio.sample_rate, audio.sample_width)

                audio_queue.put((audio_buf, False))

'''
Media Info Thread
'''

async def media_info_thread():
    sessions = await MediaManager.request_async()
    current_session = sessions.get_current_session()

    if current_session:
        if True:
            info = await current_session.try_get_media_properties_async()

            info_dict = {song_attr: info.__getattribute__(song_attr) for song_attr in dir(info) if song_attr[0] != '_'}

            info_dict['genres'] = list(info_dict['genres'])

            pbinfo = current_session.get_playback_info()
            info_dict['status'] = pbinfo.playback_status
            tlprops = current_session.get_timeline_properties()

            if tlprops.end_time != timedelta(0):
                info_dict['pos'] = tlprops.position
                info_dict['end'] = tlprops.end_time

            return info_dict
    
    else:
        raise NoMediaRunningException("No media source running")

'''
Time String
'''

def time_string(n):
    seconds = abs(int(n.seconds))

    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    if hours > 0:
        return '%d:%02d:%02d' % (hours, minutes, seconds)
    else:
        return '%d:%02d' % (minutes, seconds)

'''
Media String
'''

def media_string(udp_client):
    global textfile_first_tick, last_displayed_song
    if not textfile_first_tick:
        textfile_first_tick = True
        print(f"Now watching text file {config_music['TextFileLocation']}")

    if not os.path.exists(config_music['TextFileLocation']):
        return
    
    text = None
    with open(config_music['TextFileLocation'], 'r', encoding='utf-8') as f:
        text = f.read()

    if text is None:
        return

    duplicate_message = False
    if text == last_displayed_song:
        duplicate_message = True
        if not config_music['TextFileUpdateAlways']:
            return
    
    if text.strip() == "":
        return
    
    if not duplicate_message:
        print(f"New song detected: {text}")
    
    udp_client.send_message("/chatbox/input", [text, True, False])
    last_displayed_song = text

'''
OSC Server
'''

existing_params = {
    "Enabled": True,
    "FollowMicMute": True,
    "Pause": False,
}

class OSCServer():
    def __init__(self):
        global config
        self.dispatcher = Dispatcher()
        self.dispatcher.set_default_handler(self._def_osc_dispatch)
        self.dispatcher.map("/avatar/parameters/vrcmosc-Enabled", self.enabled)
        self.dispatcher.map("/avatar/parameters/vrcmosc-Pause", self._osc_pause)
        self.dispatcher.map("/avatar/parameters/AFK", self.AFK)

        self.server = BlockingOSCUDPServer(("127.0.0.1", config_subs['OSCControlPort']), self.dispatcher)
        self.server_thread = threading.Thread(target=self._process_osc)

    def launch(self):
        self.server_thread.start()

    def shutdown(self):
        self.server.shutdown()
        self.server_thread.join()
    

    def enabled(self, address, *args):
        print("Enabled is now: ", args[0])
        config_music['Enabled'] = args[0]
    
    def _osc_pause(self, address, *args):
        print("Pausing is now: ", args[0])
        config_subs['Pause'] = args[0]

    def AFK(self, address, *args):
        print("AFK Status: ", args[0])
        config_music['AFK'] = args[0]
        config_music['AFKSince'] = datetime.datetime.now()

    def _osc_updateconf(self, address, *args):
        key = address.split("vrcmosc-")[1]
        print("Updating config: ", key, args[0])
        
        #check configs for key
        if key in config_music:
            config_music[key] = args[0]
        elif key in config_subs:
            config_subs[key] = args[0]
    
    def _def_osc_dispatch(self, address, *args):
        #print("Recei ed unknown OSC message: ", address, args)
        pass
    
    def _process_osc(self):
        print("Starting OSC server")
        self.server.serve_forever()

class NoMediaRunningException(Exception):
    pass

'''
Music Thread

Get current song information
'''

def music_thread():
    global config, last_displayed_song, displayed_timestamp, last_reported_timestamp, client
    
    lastPaused = False
    while True:
        AFK_message = ""
        if config_music['AFK'] == True:
            AFK_message = f"AFK [ {time_string((datetime.datetime.now()-config_music['AFKSince']))} ] \u2028\u2028"

        if not config_music['Enabled']:
            time.sleep(1.5)
            continue

        if config_music['UseTextFile']:
            tick_textfile(client)
            time.sleep(1.5)
            continue

        try:
            info = asyncio.run(media_info_thread()) 
        except NoMediaRunningException:
            time.sleep(1.5)
            continue
        except Exception as e:
            print("Error getting media info: ", e, traceback.format_exc())
            time.sleep(1.5)
            continue

        song_artist, song_title = (info['artist'], info['title'])
        activity = ""

        cmd = 'WMIC PROCESS get Caption,Commandline,Processid'
        proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        for line in proc.stdout:
            line = line.decode('utf-8').strip('\r\n')

            if 'Unity.exe' in line:
                line = line.split(' ')
                line = list(filter(None, line))
                line.pop()
                
                process_name = line[0]

                if '-projectPath' in line:
                    project_path = line[line.index('-projectPath') + 1]
                    project_path = project_path.split('\\')[-1]
                    for item in line[line.index('-projectPath') + 2:]:
                        if '"' in item:
                            project_path += ' ' + item.strip('"')
                            project_path += ' ' + item.strip('""')
                        else:
                            project_path += ' ' + item
                    
                    activity = f"░▒▓ In Unity ▓▒░\u2028{project_path}\u2028\u2028"
                    break
                else:
                    activity = f"░▒▓ In Unity ▓▒░\u2028⚠winsdk_subprocess⚠\u2028-projectPath error\u2028\u2028"
                    break
            else:
                activity = config['Activity']['ActivityIdle']
            '''🥽 In PCVR'''


        current_song_string = f"\u2028{AFK_message}{activity}░▒▓ Blasting Music ▓▒░\u2028{song_artist}\u2028{song_title}\u2028░▒▓ {config['Activity']['SubText']} ▓▒░\u2028​"
        
        if len(current_song_string) >= 144 :
            current_song_string = current_song_string[:144] + "..."
        if info['status'] == GlobalSystemMediaTransportControlsSessionPlaybackStatus.PLAYING:
            '''send_to_vrc = not config_music['OnlyShowOnChange']'''
            '''if last_displayed_song != (song_artist, song_title):'''
            send_to_vrc = True
            last_displayed_song = (song_artist, song_title)
            print(f"New song detected: {song_artist} - {song_title}")

            if send_to_vrc:
                client.send_message("/chatbox/input", [current_song_string, True, False])
            lastPaused = False
        elif info['status'] == GlobalSystemMediaTransportControlsSessionPlaybackStatus.PAUSED:
            '''if lastPaused:
                time.sleep(1.5)
                continue'''
            
            client.send_message("/chatbox/input", [f"\u2028{AFK_message}{activity}░▒▓ Fuckin NOVA! ▓▒░\u2028 ⏸ {config_music['PausedFormat']}\u2028░▒▓ {config['Activity']['SubText']} ▓▒░\u2028​", True, False])
            last_displayed_song = ("", "")
            lastPaused = True
        time.sleep(1.5)
        continue


'''
Main

# output both subtitles and music
# check if music is enabled and if so, output music
# check if subtitles are enabled and if so, output subtitles
'''

def main():
    global config

    cfgfile = f"{os.path.dirname(os.path.realpath(__file__))}/Config.yml"
    print(f"Config file path: {cfgfile}")
    if os.path.isfile(cfgfile):
        with open(cfgfile) as ymlfile:
            config = load(ymlfile, Loader=Loader)
    else:
        print("Config file not found, creating one...")
        with open(cfgfile, 'w') as ymlfile:
            # dump config_music and config_subs into file
            dump(config, ymlfile, default_flow_style=False)
            # write config to files path
            ymlfile.write(f"Music:\n")
            for key in config['Music']:
                ymlfile.write(f"  {key}: {config_music[key]}\n")
            ymlfile.write(f"Subtitles:\n")
            for key in config_subs:
                ymlfile.write(f"  {key}: {config_subs[key]}\n")
            for key in config_activity:
                ymlfile.write(f"  {key}: {config_activity[key]}\n")

        print("Config file created, please edit it and restart the program.")
        sys.exit(0)
        

    #start threads
    msc = threading.Thread(target=music_thread)
    msc.start()

    '''pst = threading.Thread(target=process_audio)
    pst.start()

    cat = threading.Thread(target=audio_thread)
    cat.start()'''
    
    osc = None
    launchOSC = False

    if config_subs['FollowMicMute']:
        print("Speech will not work when muted in game.")
        launchOSC = True
    
    if config_subs['AllowOSCControl'] or config['Music']['Enabled']:
        launchOSC = True

    if launchOSC:
        osc = OSCServer()
        osc.launch()

    # join threads
    msc.join()
    '''pst.join()
    cat.join()'''

    '''if config_subs['AllowOSCControl']:
        subtitles_thread.join()'''

    if osc is not None:
        osc.shutdown()

if __name__ == "__main__":
    main()
