import os

# Force debug logs for the debug executable
os.environ["METAX_LOG_LEVEL"] = "DEBUG"

from main import main

if __name__ == "__main__":
    main()
