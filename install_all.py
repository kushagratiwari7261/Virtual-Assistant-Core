# install_all.py
import subprocess
import sys
import os

def run_command(command):
    """Run a command and print output"""
    print(f"Running: {command}")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(result.stdout)
        if result.stderr:
            print(f"Warnings: {result.stderr}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error: {e.stderr}")
        return False

def install_all():
    """Install all required modules for Alfred Assistant"""
    
    # First upgrade pip
    print("="*50)
    print("Upgrading pip...")
    print("="*50)
    run_command(f"{sys.executable} -m pip install --upgrade pip")
    
    # Install cmake first (needed for some packages)
    print("\n" + "="*50)
    print("Installing cmake...")
    print("="*50)
    run_command(f"{sys.executable} -m pip install cmake")
    
    # Basic packages
    print("\n" + "="*50)
    print("Installing basic packages...")
    print("="*50)
    basic_packages = [
        'numpy',
        'pandas',
        'requests',
        'beautifulsoup4',
        'python-dotenv',
        'pyttsx3',
        'psutil',
        'pillow',
        'plyer',
        'pywhatkit',
        'pyjokes',
        'selenium',
        'pytube'
    ]
    
    for package in basic_packages:
        print(f"\nInstalling {package}...")
        run_command(f"{sys.executable} -m pip install {package}")
    
    # Speech recognition packages
    print("\n" + "="*50)
    print("Installing speech recognition packages...")
    print("="*50)
    speech_packages = [
        'SpeechRecognition',
        'PyAudio'
    ]
    
    for package in speech_packages:
        print(f"\nInstalling {package}...")
        run_command(f"{sys.executable} -m pip install {package}")
    
    # Computer vision packages
    print("\n" + "="*50)
    print("Installing computer vision packages...")
    print("="*50)
    cv_packages = [
        'opencv-python',
        'mediapipe'
    ]
    
    for package in cv_packages:
        print(f"\nInstalling {package}...")
        run_command(f"{sys.executable} -m pip install {package}")
    
    # AI/ML packages
    print("\n" + "="*50)
    print("Installing AI/ML packages...")
    print("="*50)
    ai_packages = [
        'spacy',
        'transformers',
        'torch',
        'google-generativeai'
    ]
    
    for package in ai_packages:
        print(f"\nInstalling {package}...")
        run_command(f"{sys.executable} -m pip install {package}")
    
    # GUI/Automation packages
    print("\n" + "="*50)
    print("Installing GUI/Automation packages...")
    print("="*50)
    gui_packages = [
        'pygame',
        'pyautogui'
    ]
    
    for package in gui_packages:
        print(f"\nInstalling {package}...")
        run_command(f"{sys.executable} -m pip install {package}")
    
    # Windows specific packages
    print("\n" + "="*50)
    print("Installing Windows specific packages...")
    print("="*50)
    if os.name == 'nt':  # Windows
        run_command(f"{sys.executable} -m pip install pywin32")
    
    # Install spacy model
    print("\n" + "="*50)
    print("Installing spaCy English model...")
    print("="*50)
    run_command(f"{sys.executable} -m spacy download en_core_web_sm")
    
    # Install TensorFlow (CPU version for compatibility)
    print("\n" + "="*50)
    print("Installing TensorFlow...")
    print("="*50)
    run_command(f"{sys.executable} -m pip install tensorflow")
    
    # Verify installations
    print("\n" + "="*50)
    print("Verifying installations...")
    print("="*50)
    
    test_imports = [
        'speech_recognition',
        'pyttsx3',
        'requests',
        'bs4',
        'dotenv',
        'transformers',
        'spacy',
        'cv2',
        'mediapipe',
        'google.generativeai'
    ]
    
    for module in test_imports:
        try:
            __import__(module.replace('.', '_'))
            print(f"✓ {module} installed successfully")
        except ImportError as e:
            print(f"✗ {module} not installed: {e}")
    
    print("\n" + "="*50)
    print("Installation complete!")
    print("="*50)
    print("\nAdditional setup needed:")
    print("1. Download ChromeDriver from: https://chromedriver.chromium.org/")
    print("2. Extract and place chromedriver.exe in C:\\webdrivers\\")
    print("3. Create a .env file with your API keys:")
    print("   - NEWS_API_KEY=your_newsapi_key")
    print("   - API_KEY=your_gemini_api_key")
    print("\nRun your assistant with: python alfred_assistant.py")

if __name__ == "__main__":
    print("Alfred Assistant - Complete Installation Script")
    print("This will install all required packages for the assistant.")
    print("The installation may take 10-15 minutes depending on your internet speed.")
    
    response = input("\nDo you want to proceed with installation? (y/n): ")
    if response.lower() == 'y':
        install_all()
    else:
        print("Installation cancelled.")