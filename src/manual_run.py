import asyncio
from server import get_presentation_rules, generate_presentation
import json
import os


width = 115

def print_menu():
    print(r"""
  _____                    _       _           _   ____  ____ _____    ____                           _             
 |_   _|__ _ __ ___  _ __ | | __ _| |_ ___  __| | |  _ \|  _ \_   _|  / ___| ___ _ __   ___ _ __ __ _| |_ ___  _ __ 
   | |/ _ \ '_ ` _ \| '_ \| |/ _` | __/ _ \/ _` | | |_) | |_) || |   | |  _ / _ \ '_ \ / _ \ '__/ _` | __/ _ \| '__|
   | |  __/ | | | | | |_) | | (_| | ||  __/ (_| | |  __/|  __/ | |   | |_| |  __/ | | |  __/ | | (_| | || (_) | |   
   |_|\___|_| |_| |_| .__/|_|\__,_|\__\___|\__,_| |_|   |_|    |_|    \____|\___|_| |_|\___|_|  \__,_|\__\___/|_|   
                    |_|                                                                                             

    """)
    print("\n")
    print("="*width)
    print("Select the action you want:")
    print("-"*width)
    print("1. Generate Prompt Text, then you can use any AI you like to generate the json content.")
    print("2. Generate Presentation from JSON file you generated use above prompt.")
    print("0. Exit")
    print("="*width)

def pretty_print_prompt(prompt_data):
    title = "Copy below Prompt"
    print("\n"*3)
    print("="*int((width-len(title))/2),title,"="*int((width-len(title))/2))
    if isinstance(prompt_data, dict) or isinstance(prompt_data, list):
        print(json.dumps(prompt_data, indent=4))
    else:
        print(prompt_data)
    print("="*width,"\n"*3)

def generate_prompt():
    try:
        result = asyncio.run(get_presentation_rules(ctx=None))
        pretty_print_prompt(result)
    except Exception as e:
        print(f"[Error] Failed to generate prompt: {e}")

def generate_from_file():
    json_file_path = input("Enter the path to JSON file. For example: D:/MCP/playground/presentation.json\nFile path:  ").strip()
    if not os.path.exists(json_file_path):
        print(f"[Error] File does not exist: {json_file_path}")
        return
    try:
        filename = os.path.basename(json_file_path)
        os.environ["WORK_DIR"] = os.path.dirname(json_file_path)
        message = asyncio.run(generate_presentation(ctx=None, json_content_filename = filename))
        print("\n"*2)
        print("[Success] Presentation generated successfully.")
        print(f"Presentation saved to: {message['save_location']}")
        print(f"Presentation filename: {message['filename']}")
        print(f"Slides count: {message['slides_count']}")
    except Exception as e:
        print(f"[Error] Failed to generate presentation:\n {e}")

def main():
    while True:
        print_menu()
        choice = input("Choose an option (0/1/2): ").strip()
        if choice == "1":
            generate_prompt()

        elif choice == "2":
            generate_from_file()
        elif choice == "0":
            print("Exiting. Goodbye!")
        else:
            print("[Warning] Invalid choice. Please enter 0, 1, or 2.")
        
        break

if __name__ == "__main__":
    main()
