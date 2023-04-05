import winsound
import keyboard, pyttsx3, os
from threading import Thread, Condition
from common import static_class, ReadFromClipboard

# Hide pygame welcome message.
os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"
from pygame import mixer

@static_class
class TTS_House:
    # Initializing a queue for holding incoming text.
    textQueue = set()
    curr_text = ""
    
    # Initializing the TTS enginge.
    engine = pyttsx3.init()
    
    # Configuring the volume and rate of the speech.
    engine.setProperty('volume', 1.0)
    engine.setProperty('rate', 250)
    
    # `status` variable for stoping or pausing the speech.
    status = 1 #  1:play | 2:pause | 3:stopped
    
    # For not triggering the operations multiple times.
    op_called = True
    
    # Initializing pygame mixer.
    mixer.init()
    
    condition = Condition()
    
    @staticmethod
    def ScheduleSpeak():
        # os.system(r'"c:\Program Files\Ditto\Ditto.exe" /disconnect') # Disconnect Ditto from the clipboard to prevent the copied text from being saved.
        keyboard.send("ctrl+c")
        # os.system(r'"c:\Program Files\Ditto\Ditto.exe" /connect')    # Reconnect Ditto to the clipboard.
        
        # Read copied text from clipboard
        text = ReadFromClipboard()
        if not text:
            winsound.PlaySound(r"SFX\record-scratch-2.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            return
        
        text = text.replace("\n", " ")
        
        # Add the copied text to the text queue
        if text != TTS_House.curr_text:
            TTS_House.textQueue.add(text)
        
        # Not busy means that no threads are running so create one.
        if len(TTS_House.textQueue) and not mixer.music.get_busy():
            Thread(target=TTS_House.SpeakAllTextFromQueue).start()
    
    @staticmethod
    def SpeakAllTextFromQueue():
        while True:
            # Get the next item from the queue
            TTS_House.curr_text = TTS_House.textQueue.pop()
            
            # Save the text to a temporary file
            if os.path.exists(os.path.join("SFX", "tts_temp.wav")):    # temp exists, temp1 doesn't.
                outFileName = os.path.join("SFX", "tts_temp1.wav")
            
            elif os.path.exists(os.path.join("SFX", "tts_temp1.wav")): # temp1 exists, temp doesn't .
                outFileName = os.path.join("SFX", "tts_temp.wav")
            
            else:                                                      # Both files exist (only at the start of the program)
                os.remove(os.path.join("SFX", "tts_temp.wav"))
                os.remove(os.path.join("SFX", "tts_temp1.wav"))
                outFileName = os.path.join("SFX", "tts_temp.wav")
            
            TTS_House.engine.save_to_file(TTS_House.curr_text, outFileName)
            TTS_House.engine.runAndWait()
            
            # Load and play the temporary file using pygame mixer.
            mixer.music.load(outFileName)
            mixer.music.play()
            TTS_House.status = 1
            TTS_House.op_called = True
            
            # Delete other temp file if it does exist.
            if outFileName[-5] == "1":
                os.remove(os.path.join("SFX", "tts_temp.wav"))
            elif os.path.exists(os.path.join("SFX", "tts_temp1.wav")):
                os.remove(os.path.join("SFX", "tts_temp1.wav"))
            
            # Wait until the audio finishes playing or the queue is empty
            while True:
                if not TTS_House.op_called:
                    match TTS_House.status:
                        case 2:
                            mixer.music.pause()
                            with TTS_House.condition:
                                TTS_House.condition.wait()
                                mixer.music.unpause()
                                TTS_House.status = 1
                        case 3:
                            mixer.music.stop()
                            TTS_House.textQueue.clear()
                            break
                    TTS_House.op_called = True
                elif not mixer.music.get_busy():
                    break
            
            # If the queue is empty, stop the speaking thread
            if not len(TTS_House.textQueue):
                break
        TTS_House.curr_text = ""


if __name__ == "__main__":
    import time
    import pyaudio
    import wave
    
    # Source: https://stackoverflow.com/a/69042516/14682148
    class SamplePlayer:
        def __init__(self, filepath=""):
            self.file = filepath
            self.paused = True
            self.playing = False
            
            self.audio_length = 0
            self.current_sec = 0
        
        def start_playing(self):
            with wave.open(self.file, "rb") as wf:
                # Another way for querying the audio file length: https://stackoverflow.com/questions/7833807/get-wav-file-length-or-duration
                # import librosa; librosa.get_duration(filename='my.wav')
                self.audio_length = wf.getnframes() / float(wf.getframerate())
                if self.audio_length < 5:
                    winsound.PlaySound(self.file, winsound.SND_FILENAME|winsound.SND_ASYNC)
                    return
                else:
                    p = pyaudio.PyAudio()
                    chunk = 1024
                    stream = p.open(format=p.get_format_from_width(wf.getsampwidth()),
                                    channels=wf.getnchannels(),
                                    rate=wf.getframerate(),
                                    output=True)
                    data = wf.readframes(chunk)
                    
                    chunk_total = 0
                    while data != b"" and self.playing:
                        if self.paused:
                            time.sleep(0.1)
                        else:
                            chunk_total += chunk
                            stream.write(data)
                            data = wf.readframes(chunk)
                            self.current_sec = chunk_total/wf.getframerate()
                    
                    self.playing = False
                    stream.close()   
                    p.terminate()
        
        def pause(self):
            self.paused = True
        
        def play(self):
            if not self.playing:
                self.playing = True
                Thread(target=self.start_playing, daemon=True).start()
            
            if self.paused:
                self.paused = False
        
        def stop(self):
            self.playing = False
