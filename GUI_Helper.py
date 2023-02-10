import keyboard, winsound, pyttsx3, os
from threading import Thread, Condition
from common import Singleton, ReadFromClipboard

# Hide pygame welcome message.
os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"
from pygame import mixer

class TTS_House(metaclass=Singleton):
    def __init__(self):
        # Initializing a queue for holding incoming text.
        self.textQueue = set()
        self.curr_text = ""
        
        # Initializing the TTS enginge.
        self.engine = pyttsx3.init()
        
        # Configuring the volume and rate of the speech.
        self.engine.setProperty('volume', 1.0)
        self.engine.setProperty('rate', 250)
        
        # `status` variable for stoping or pausing the speech.
        self.status = 1 #  1:play | 2:pause | 3:stopped
        
        # For not triggering the operations multiple times.
        self.op_called = True
        
        # Initializing pygame mixer.
        mixer.init()
        
        self.condition = Condition()
    
    def ScheduleSpeak(self):
        keyboard.press("ctrl")
        keyboard.press("c")
        keyboard.release("c")
        keyboard.release("ctrl")
        
        # Read copied text from clipboard
        text = ReadFromClipboard()
        if not text:
            winsound.PlaySound(r"SFX\record-scratch-2.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            return
        
        # Add the copied text to the text queue
        if text != self.curr_text:
            self.textQueue.add(text)
        
        # Not busy means that no threads are running so create one.
        if not mixer.music.get_busy():
            Thread(target=self.SpeakAllTextFromQueue).start()
    
    def SpeakAllTextFromQueue(self):
        while True:
            # Get the next item from the queue
            self.curr_text = self.textQueue.pop()
            
            # Save the text to a temporary file
            if os.path.exists(os.path.join("SFX", "TTStemp.wav")):
                outFileName = os.path.join("SFX", "TTStemp1.wav")
            else:
                outFileName = os.path.join("SFX", "TTStemp.wav")
            
            self.engine.save_to_file(self.curr_text, outFileName)
            self.engine.runAndWait()
            
            # Load and play the temporary file using pygame mixer.
            mixer.music.load(outFileName)
            mixer.music.play()
            self.status = 1
            self.op_called = True
            
            # Delete other temp file if it does exist.
            if outFileName[4] == "1":
                os.remove(os.path.join("SFX", "TTStemp.wav"))
            elif os.path.exists(os.path.join("SFX", "TTStemp1.wav")):
                os.remove(os.path.join("SFX", "TTStemp1.wav"))
            
            # Wait until the audio finishes playing or the queue is empty
            while True:
                if not self.op_called:
                    match self.status:
                        case 2:
                            mixer.music.pause()
                            with self.condition:
                                self.condition.wait()
                                mixer.music.unpause()
                                self.status = 1
                        case 3:
                            mixer.music.stop()
                            self.textQueue.clear()
                            break
                    self.op_called = True
                elif not mixer.music.get_busy():
                    break
            
            # If the queue is empty, stop the speaking thread
            if not len(self.textQueue):
                break
        self.curr_text = ""


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
