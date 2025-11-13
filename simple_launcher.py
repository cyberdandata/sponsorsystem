import os
import sys
import subprocess
import time
import webbrowser

def main():
    print("🎯 Starting Sponsorship Management System...")
    
    # ALWAYS use the directory where the original files are located
    # This ensures the .exe works from the root directory, not dist folder
    if getattr(sys, 'frozen', False):
        # If running as .exe, we need to find the root directory
        # The .exe might be in dist, but we need to go up to root
        exe_dir = os.path.dirname(sys.executable)
        
        # Check if we're in dist folder and need to go up one level
        if os.path.basename(exe_dir) == 'dist':
            base_dir = os.path.dirname(exe_dir)  # Go up to root directory
        else:
            base_dir = exe_dir  # Already in root directory
    else:
        # Running as script - use the script directory (root)
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    print(f"📁 Root directory: {base_dir}")
    
    # Change to the root directory where server.js is located
    os.chdir(base_dir)
    
    # Verify all required files exist in root directory
    required_files = {
        'server.js': os.path.join(base_dir, 'server.js'),
        'package.json': os.path.join(base_dir, 'package.json'),
        'public/index.html': os.path.join(base_dir, 'public', 'index.html')
    }
    
    for file_name, file_path in required_files.items():
        if not os.path.exists(file_path):
            print(f"❌ Missing: {file_name} at {file_path}")
            print("Please make sure all files are in the correct location")
            input("Press Enter to exit...")
            return
        else:
            print(f"✅ Found: {file_name}")
    
    # Check Node.js
    try:
        result = subprocess.run(['node', '--version'], capture_output=True, text=True, check=True)
        print(f"✅ Node.js: {result.stdout.strip()}")
    except:
        print("❌ Node.js not found. Please install Node.js")
        input("Press Enter to exit...")
        return
    
    # Install dependencies
    print("📦 Installing dependencies...")
    try:
        subprocess.run(['npm', 'install'], check=True, cwd=base_dir)
        print("✅ Dependencies installed")
    except:
        print("⚠️  Dependency installation failed, continuing...")
    
    # Open browser
    print("🌍 Opening browser...")
    webbrowser.open('http://localhost:3000')
    
    # Start server
    print("🚀 Starting server...")
    print("📍 Access at: http://localhost:3000")
    print("🛑 Press Ctrl+C to stop the server")
    print("=" * 50)
    
    try:
        # Start server from the root directory
        subprocess.run(['node', 'server.js'], check=True, cwd=base_dir)
    except KeyboardInterrupt:
        print("\n👋 Server stopped")
    except Exception as e:
        print(f"❌ Server error: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()