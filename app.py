import os
import sys
import time
import random
import json
import re
import subprocess
import threading
from datetime import datetime
from pathlib import Path
import webbrowser
import pyautogui
import pyperclip
import psutil
import speech_recognition as sr
import requests
from plyer import notification
import pyjokes
import win32com.client
import urllib.parse
import pygetwindow as gw

# ========== CONFIGURATION ==========
class Config:
    NAME = "Alfred"
    VERSION = "Ultimate 4.0"
    
    # Paths
    APPS_PATH = {
        'notepad': 'notepad.exe',
        'calculator': 'calc.exe',
        'browser': 'chrome.exe',
        'explorer': 'explorer.exe',
        'paint': 'mspaint.exe',
        'cmd': 'cmd.exe',
        'word': 'winword.exe',
        'excel': 'excel.exe',
        'powerpoint': 'powerpnt.exe',
        'vscode': 'code.exe',
        'spotify': 'spotify.exe',
        'whatsapp': 'whatsapp://',
        'telegram': 'tg://',
        'vlc': 'vlc.exe',
        'obs': 'obs64.exe',
        'discord': 'discord://',
        'steam': 'steam://',
        'chrome': 'chrome.exe',
        'edge': 'msedge.exe',
        'firefox': 'firefox.exe',
        'photoshop': 'photoshop.exe',
        'premiere': 'Adobe Premiere Pro.exe',
        'illustrator': 'illustrator.exe',
    }
    
    # URLs
    URLS = {
        'youtube': 'https://www.youtube.com',
        'google': 'https://www.google.com',
        'gmail': 'https://mail.google.com',
        'github': 'https://github.com',
        'wikipedia': 'https://wikipedia.org',
        'netflix': 'https://netflix.com',
        'prime': 'https://primevideo.com',
        'hotstar': 'https://hotstar.com',
        'spotify web': 'https://open.spotify.com',
        'chatgpt': 'https://chat.openai.com',
        'gemini': 'https://gemini.google.com',
        'drive': 'https://drive.google.com',
        'maps': 'https://maps.google.com',
        'facebook': 'https://facebook.com',
        'instagram': 'https://instagram.com',
        'twitter': 'https://twitter.com',
        'linkedin': 'https://linkedin.com',
        'reddit': 'https://reddit.com',
    }
    
    # File paths
    SCREENSHOTS_DIR = "screenshots"
    LOGS_DIR = "logs"

# ========== UTILITY FUNCTIONS ==========
class Utils:
    @staticmethod
    def ensure_directories():
        """Create necessary directories"""
        for dir_path in [Config.SCREENSHOTS_DIR, Config.LOGS_DIR]:
            Path(dir_path).mkdir(exist_ok=True)
    
    @staticmethod
    def log(message: str, level: str = "INFO"):
        """Log messages"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_msg = f"[{timestamp}] {message}"
        print(f"📝 {log_msg}")
    
    @staticmethod
    def get_timestamp() -> str:
        """Get current timestamp"""
        return datetime.now().strftime("%Y%m%d_%H%M%S")
    
    @staticmethod
    def run_command(command: str, wait: bool = True) -> bool:
        """Run system command"""
        try:
            if wait:
                subprocess.run(command, shell=True, check=True)
            else:
                subprocess.Popen(command, shell=True)
            return True
        except Exception as e:
            Utils.log(f"Command failed: {e}", "ERROR")
            return False

# ========== SPEECH ENGINE ==========
class SpeechEngine:
    def __init__(self):
        self.speaker = None
        self._init_speaker()
    
    def _init_speaker(self):
        """Initialize Windows speech engine"""
        try:
            self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
            self.speaker.Rate = 0
            self.speaker.Volume = 100
            Utils.log("Speech engine initialized")
        except Exception as e:
            Utils.log(f"Speech init failed: {e}", "WARNING")
            self.speaker = None
    
    def speak(self, text: str):
        """Speak text using Windows SAPI"""
        print(f"\n🤖 {Config.NAME}: {text}")
        
        if self.speaker:
            try:
                self.speaker.Speak(text)
                return True
            except Exception as e:
                Utils.log(f"Speech failed: {e}", "WARNING")
        
        # Fallback: Use PowerShell speech
        try:
            safe_text = text.replace('"', '\\"').replace("'", "\\'")
            ps_command = f'powershell -Command "Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak(\'{safe_text}\')"'
            subprocess.run(ps_command, shell=True, capture_output=True)
            return True
        except:
            pass
        
        return False

# ========== VOICE RECOGNITION ==========
class VoiceRecognition:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.microphone = None
    
    def listen(self) -> str:
        """Listen for voice input - IMPROVED VERSION"""
        try:
            if self.microphone is None:
                self.microphone = sr.Microphone()
            
            with self.microphone as source:
                # Reduce ambient noise sensitivity
                self.recognizer.energy_threshold = 300
                self.recognizer.dynamic_energy_threshold = True
                
                print("\n🎤 Listening... (Speak clearly)")
                
                # Listen with better parameters
                audio = self.recognizer.listen(
                    source, 
                    timeout=3,
                    phrase_time_limit=5
                )
                
                # Use Google Speech Recognition
                text = self.recognizer.recognize_google(audio, language="en-IN")
                print(f"👤 You said: {text}")
                return text.lower()
                
        except sr.UnknownValueError:
            print("❌ Could not understand audio")
            return ""
        except sr.RequestError as e:
            print(f"❌ Speech service error: {e}")
            return ""
        except Exception as e:
            print(f"❌ Error: {e}")
            return ""
    
    def listen_with_retry(self, retries=3):
        """Listen with retries"""
        for i in range(retries):
            result = self.listen()
            if result:
                return result
            time.sleep(0.5)
        return ""

# ========== AUTOMATION ENGINE ==========
class AutomationEngine:
    @staticmethod
    def open_application(app_name: str) -> str:
        """Open any application"""
        app_lower = app_name.lower()
        
        # Check in apps
        for key, path in Config.APPS_PATH.items():
            if key in app_lower:
                try:
                    if '://' in path:  # URL scheme
                        webbrowser.open(path)
                    elif os.path.exists(path):
                        os.startfile(path)
                    else:
                        os.system(f'start {path}')
                    return f"Opening {key}..."
                except Exception as e:
                    return f"Failed to open {key}"
        
        # Check in URLs
        for key, url in Config.URLS.items():
            if key in app_lower:
                webbrowser.open(url)
                return f"Opening {key}..."
        
        # Try direct
        try:
            os.system(f'start {app_name}')
            return f"Trying to open {app_name}..."
        except:
            return f"Don't know how to open {app_name}"

    @staticmethod
    def close_application(app_name: str) -> str:
        """Close application"""
        app_lower = app_name.lower()
        
        process_map = {
            'chrome': 'chrome.exe',
            'browser': 'chrome.exe',
            'notepad': 'notepad.exe',
            'calculator': 'calc.exe',
            'explorer': 'explorer.exe',
            'word': 'winword.exe',
            'excel': 'excel.exe',
            'powerpoint': 'powerpnt.exe',
            'spotify': 'spotify.exe',
            'vlc': 'vlc.exe',
            'steam': 'steam.exe',
            'discord': 'discord.exe',
            'obs': 'obs64.exe'
        }
        
        for key, process in process_map.items():
            if key in app_lower:
                try:
                    os.system(f'taskkill /f /im {process} 2>nul')
                    return f"Closing {key}..."
                except:
                    return f"Failed to close {key}"
        
        return f"Don't know how to close {app_name}"

    @staticmethod
    def youtube_search(query: str) -> str:
        """Search and play YouTube"""
        try:
            search_url = f"https://www.youtube.com/results?search_query={urllib.parse.quote(query)}"
            webbrowser.open(search_url)
            
            # Auto-play first result after delay
            def auto_play():
                time.sleep(4)
                pyautogui.press('tab', presses=3)
                time.sleep(0.5)
                pyautogui.press('enter')
            
            threading.Thread(target=auto_play, daemon=True).start()
            return f"Searching YouTube for {query} and playing first result..."
        except Exception as e:
            return f"Failed: {str(e)}"

    @staticmethod
    def youtube_control(action: str) -> str:
        """Control YouTube playback"""
        actions = {
            'play': 'k',
            'pause': 'k',
            'next': 'shift+n',
            'previous': 'shift+p',
            'fullscreen': 'f',
            'mute': 'm',
            'skip forward': 'l',
            'skip backward': 'j'
        }
        
        if action in actions:
            pyautogui.press(actions[action])
            return f"YouTube {action}ed"
        return f"Unknown action: {action}"

    @staticmethod
    def take_screenshot() -> str:
        """Take screenshot"""
        try:
            filename = f"screenshot_{Utils.get_timestamp()}.png"
            path = Path(Config.SCREENSHOTS_DIR) / filename
            pyautogui.screenshot().save(path)
            return f"Screenshot saved: {filename}"
        except Exception as e:
            return f"Failed: {str(e)}"

    @staticmethod
    def send_whatsapp(phone: str, message: str = "Hello from Alfred!") -> str:
        """Send WhatsApp message"""
        try:
            phone_clean = re.sub(r'\D', '', phone)
            if len(phone_clean) < 10:
                return "Invalid phone number"
            
            url = f"whatsapp://send?phone={phone_clean}&text={urllib.parse.quote(message)}"
            webbrowser.open(url)
            
            # Auto-send after delay
            def auto_send():
                time.sleep(5)
                pyautogui.press('enter')
            
            threading.Thread(target=auto_send, daemon=True).start()
            return f"Sending WhatsApp to {phone}: {message}"
        except Exception as e:
            return f"Failed: {str(e)}"

    @staticmethod
    def send_email(to: str, subject: str, body: str) -> str:
        """Send email"""
        try:
            url = f"mailto:{to}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
            webbrowser.open(url)
            return f"Preparing email to {to}: {subject}"
        except Exception as e:
            return f"Failed: {str(e)}"

    @staticmethod
    def web_search(query: str) -> str:
        """Search web"""
        try:
            url = f"https://www.google.com/search?q={urllib.parse.quote(query)}"
            webbrowser.open(url)
            return f"Searching for: {query}"
        except:
            return "Search failed"

    @staticmethod
    def system_info() -> str:
        """Get system info"""
        try:
            cpu = psutil.cpu_percent()
            memory = psutil.virtual_memory()
            return f"CPU: {cpu}%, Memory: {memory.percent}% used"
        except:
            return "System info unavailable"

    @staticmethod
    def control_volume(action: str) -> str:
        """Control volume"""
        try:
            if action == "up":
                for _ in range(5):
                    pyautogui.press('volumeup')
                return "Volume increased"
            elif action == "down":
                for _ in range(5):
                    pyautogui.press('volumedown')
                return "Volume decreased"
            elif action == "mute":
                pyautogui.press('volumemute')
                return "Volume muted"
            else:
                return "Unknown volume command"
        except:
            return "Volume control failed"

    @staticmethod
    def create_file(filename: str, content: str = "") -> str:
        """Create file"""
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            return f"Created file: {filename}"
        except Exception as e:
            return f"Failed: {str(e)}"

    @staticmethod
    def read_file(filename: str) -> str:
        """Read file"""
        try:
            if os.path.exists(filename):
                with open(filename, 'r', encoding='utf-8') as f:
                    content = f.read()
                return f"File {filename}:\n{content[:200]}..."
            return f"File not found: {filename}"
        except Exception as e:
            return f"Failed: {str(e)}"

    @staticmethod
    def execute_command(cmd: str) -> str:
        """Execute system command"""
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            output = result.stdout if result.stdout else result.stderr
            return f"Command output:\n{output[:300]}"
        except Exception as e:
            return f"Command failed: {str(e)}"

    @staticmethod
    def set_reminder(text: str, minutes: int = 5) -> str:
        """Set reminder"""
        def reminder():
            time.sleep(minutes * 60)
            notification.notify(
                title="⏰ Alfred Reminder",
                message=text,
                timeout=10
            )
        
        threading.Thread(target=reminder, daemon=True).start()
        return f"Reminder set for {minutes} minutes: {text}"

    @staticmethod
    def type_text(text: str) -> str:
        """Type text"""
        try:
            pyautogui.write(text, interval=0.1)
            return f"Typed: {text[:50]}..."
        except:
            return "Typing failed"

    @staticmethod
    def press_key(key: str) -> str:
        """Press key"""
        try:
            pyautogui.press(key)
            return f"Pressed {key}"
        except:
            return f"Failed to press {key}"

# ========== COMMAND PROCESSOR ==========
class CommandProcessor:
    def __init__(self):
        self.automation = AutomationEngine()
        self.commands = self._load_commands()
    
    def _load_commands(self):
        """Load command patterns"""
        return {
            # Open commands
            r'open (.+)': lambda m: self.automation.open_application(m.group(1)),
            r'start (.+)': lambda m: self.automation.open_application(m.group(1)),
            r'launch (.+)': lambda m: self.automation.open_application(m.group(1)),
            
            # Close commands
            r'close (.+)': lambda m: self.automation.close_application(m.group(1)),
            r'quit (.+)': lambda m: self.automation.close_application(m.group(1)),
            r'exit (.+)': lambda m: self.automation.close_application(m.group(1)),
            
            # YouTube
            r'play (.+) on youtube': lambda m: self.automation.youtube_search(m.group(1)),
            r'search (.+) on youtube': lambda m: self.automation.youtube_search(m.group(1)),
            r'youtube search (.+)': lambda m: self.automation.youtube_search(m.group(1)),
            r'youtube (.+)': lambda m: self.automation.youtube_search(m.group(1)),
            r'pause youtube': lambda: self.automation.youtube_control('pause'),
            r'play youtube': lambda: self.automation.youtube_control('play'),
            r'next video': lambda: self.automation.youtube_control('next'),
            r'fullscreen youtube': lambda: self.automation.youtube_control('fullscreen'),
            
            # WhatsApp
            r'send whatsapp to (\d+) (.+)': lambda m: self.automation.send_whatsapp(m.group(1), m.group(2)),
            r'send whatsapp (.+) to (\d+)': lambda m: self.automation.send_whatsapp(m.group(2), m.group(1)),
            r'whatsapp (.+) to (\d+)': lambda m: self.automation.send_whatsapp(m.group(2), m.group(1)),
            
            # Screenshot
            r'take screenshot': lambda: self.automation.take_screenshot(),
            r'capture screen': lambda: self.automation.take_screenshot(),
            r'screenshot': lambda: self.automation.take_screenshot(),
            
            # Volume
            r'volume up': lambda: self.automation.control_volume('up'),
            r'volume down': lambda: self.automation.control_volume('down'),
            r'increase volume': lambda: self.automation.control_volume('up'),
            r'decrease volume': lambda: self.automation.control_volume('down'),
            r'mute': lambda: self.automation.control_volume('mute'),
            
            # Search
            r'search (.+)': lambda m: self.automation.web_search(m.group(1)),
            r'google (.+)': lambda m: self.automation.web_search(m.group(1)),
            
            # System
            r'system info': lambda: self.automation.system_info(),
            r'computer info': lambda: self.automation.system_info(),
            r'system status': lambda: self.automation.system_info(),
            
            # Files
            r'create file (.+)': lambda m: self.automation.create_file(m.group(1)),
            r'read file (.+)': lambda m: self.automation.read_file(m.group(1)),
            
            # Commands
            r'run command (.+)': lambda m: self.automation.execute_command(m.group(1)),
            r'execute (.+)': lambda m: self.automation.execute_command(m.group(1)),
            
            # Reminders
            r'remind me to (.+) in (\d+) minutes': lambda m: self.automation.set_reminder(m.group(1), int(m.group(2))),
            r'set reminder (.+) in (\d+) minutes': lambda m: self.automation.set_reminder(m.group(1), int(m.group(2))),
            r'reminder (.+) in (\d+) minutes': lambda m: self.automation.set_reminder(m.group(1), int(m.group(2))),
            
            # Typing
            r'type (.+)': lambda m: self.automation.type_text(m.group(1)),
            r'press (.+)': lambda m: self.automation.press_key(m.group(1)),
            
            # Email
            r'send email (.+) to (.+)': lambda m: self.automation.send_email(m.group(2), "Message from Alfred", m.group(1)),
        }
    
    def process(self, command: str) -> str:
        """Process command and return response"""
        if not command or command == "":
            return "I didn't hear anything"
        
        print(f"⚡ Processing: {command}")
        
        # Check for exit
        if any(word in command for word in ['exit', 'quit', 'goodbye', 'bye', 'stop']):
            return "exit"
        
        # Check for greetings
        if any(word in command for word in ['hello', 'hi', 'hey']):
            return random.choice([
                "Hello! I'm Alfred, your automation assistant.",
                "Hi there! Ready to help with automation.",
                "Hey! What can I automate for you today?"
            ])
        
        # Check for thanks
        if any(word in command for word in ['thank', 'thanks']):
            return random.choice([
                "You're welcome!",
                "Happy to help!",
                "My pleasure!"
            ])
        
        # Check for how are you
        if 'how are you' in command:
            return random.choice([
                "I'm functioning perfectly!",
                "All systems operational!",
                "Ready for automation tasks!"
            ])
        
        # Check for jokes
        if 'joke' in command:
            try:
                joke = pyjokes.get_joke()
                return joke
            except:
                jokes = [
                    "Why don't scientists trust atoms? Because they make up everything!",
                    "Why did the math book look sad? Because it had too many problems."
                ]
                return random.choice(jokes)
        
        # Check for time
        if 'time' in command:
            return f"The time is {datetime.now().strftime('%I:%M %p')}"
        
        # Check for date
        if 'date' in command:
            return f"Today is {datetime.now().strftime('%B %d, %Y')}"
        
        # Check for capabilities
        if any(phrase in command for phrase in ['what can you do', 'capabilities', 'help']):
            return """I can automate:
• Open/close applications
• Control YouTube: play, search, pause, next
• Send WhatsApp messages
• Take screenshots
• Control volume
• Search the web
• Create/read files
• Execute commands
• Set reminders
• Type text
• Send emails
• And much more!"""
        
        # Check for Jarvis mode
        if any(word in command for word in ['jarvis', 'iron man', 'behave like']):
            return "Activating J.A.R.V.I.S. protocol. Just Another Rather Very Intelligent System online. At your service, sir."
        
        # Try to match command patterns
        for pattern, action in self.commands.items():
            match = re.match(pattern, command, re.IGNORECASE)
            if match:
                try:
                    if match.groups():
                        return action(match)
                    else:
                        return action()
                except Exception as e:
                    return f"Error executing command: {str(e)}"
        
        # Default response
        return "I can help with automation. Try: 'open chrome', 'play music on youtube', 'take screenshot', or 'send whatsapp to 1234567890 hello'"

# ========== MAIN ASSISTANT ==========
class Alfred:
    def __init__(self):
        Utils.ensure_directories()
        self.speech = SpeechEngine()
        self.voice = VoiceRecognition()
        self.processor = CommandProcessor()
        self.running = False
        
        Utils.log(f"{Config.NAME} v{Config.VERSION} initialized")
    
    def start(self):
        """Start the assistant"""
        self.running = True
        
        print("\n" + "="*70)
        print(f"🤖 {Config.NAME} - ULTIMATE AUTOMATION ASSISTANT v{Config.VERSION}")
        print("="*70)
        print("🚀 FULL SYSTEM AUTOMATION READY")
        print("\n📋 COMMAND EXAMPLES:")
        print("• 'open chrome' - Open applications")
        print("• 'play music on youtube' - Search & play YouTube")
        print("• 'send whatsapp to 1234567890 hello there' - Send WhatsApp")
        print("• 'take screenshot' - Capture screen")
        print("• 'volume up' / 'volume down' / 'mute' - Control volume")
        print("• 'search python tutorials' - Web search")
        print("• 'system info' - Get system status")
        print("• 'create file notes.txt' - Create files")
        print("• 'remind me to call mom in 10 minutes' - Set reminders")
        print("• 'type hello world' - Type text")
        print("• 'press enter' - Press keys")
        print("• 'behave like Jarvis' - Activate Iron Man mode")
        print("• 'exit' - Quit Alfred")
        print("="*70)
        
        welcome = random.choice([
            f"{Config.NAME} online. All automation systems ready.",
            "Automation engine initialized. Ready for commands.",
            "Systems operational. What shall we automate today?"
        ])
        self.speech.speak(welcome)
    
    def run(self):
        """Main run loop"""
        while self.running:
            try:
                print("\n" + "-"*50)
                print("🎤 Speak your automation command...")
                print("(Press Ctrl+C to exit)")
                print("-"*50)
                
                # Get voice input with retry
                command = self.voice.listen_with_retry()
                
                if command:
                    response = self.processor.process(command)
                    
                    if response == "exit":
                        farewell = random.choice([
                            "Goodbye! Automation complete.",
                            "Shutting down. Until next time!",
                            "Alfred signing off. Have a great day!"
                        ])
                        self.speech.speak(farewell)
                        self.running = False
                        break
                    
                    self.speech.speak(response)
                else:
                    # No command heard
                    if random.random() < 0.3:
                        prompt = random.choice([
                            "I'm listening for automation commands...",
                            "Ready for your next command.",
                            "Awaiting automation instructions."
                        ])
                        self.speech.speak(prompt)
                
            except KeyboardInterrupt:
                print("\n\n🛑 Stopping Alfred...")
                self.speech.speak("Shutting down automation systems.")
                self.running = False
                break
            except Exception as e:
                Utils.log(f"Error: {e}", "ERROR")
                time.sleep(1)
        
        print("\n" + "="*70)
        print("👋 Alfred automation assistant stopped")
        print("="*70)

# ========== MAIN EXECUTION ==========
if __name__ == "__main__":
    print("🚀 Initializing Alfred Ultimate Automation Assistant...")
    
    # Check and install missing packages
    required = ['pyautogui', 'pyjokes', 'plyer', 'requests', 'pywin32', 'pygetwindow']
    for package in required:
        try:
            __import__(package.replace('-', '_'))
        except ImportError:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    
    # Create and run assistant
    assistant = Alfred()
    
    try:
        assistant.start()
        assistant.run()
    except Exception as e:
        print(f"💥 Fatal error: {e}")
        input("Press Enter to exit...")