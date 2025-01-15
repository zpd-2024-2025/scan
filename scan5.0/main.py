import cv2
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os
import time

# Google Drive setup using service account
SERVICE_ACCOUNT_FILE = 'service_account.json'
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Authenticate with the Google Drive API using a service account
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)

parent_folder_file_id = '1wgZ8hyW8KjOf9xkmIWzMLxuppaIRZR3L'

def initialize_drive():
    """Initialize Google Drive and return the Drive object and photo folder ID."""

    # Define parent folder ID where photos folder will be created
    parent_folder_id = '1sUKn3qGY-_ISWG_YfOn9ThDy1B8kDMLG'
    # Generate a folder name based on the current date
    photo_folder_name = f'photos_{datetime.now().strftime("%Y-%m-%d")}'
    
    # Query to check if the folder already exists in the parent folder
    query = f"name='{photo_folder_name}' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
    folder_list = drive_service.files().list(q=query).execute()

    # If the folder exists, get its ID, otherwise create it
    if folder_list['files']:
        photo_folder_id = folder_list['files'][0]['id']
    else:
        folder_metadata = {
            'name': photo_folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }
        folder = drive_service.files().create(body=folder_metadata).execute()
        photo_folder_id = folder['id']

    return drive_service, photo_folder_id

def save_photo_to_drive(drive, folder_id, frame, photo_name):
    """Save a photo to Google Drive."""

    # Create a local folder to temporarily store the photo
    local_photo_folder = 'photo'
    os.makedirs(local_photo_folder, exist_ok=True)
    photo_path = os.path.join(local_photo_folder, photo_name)
    
    # Save the frame to a local file
    cv2.imwrite(photo_path, frame)

    # Upload the photo to Google Drive
    file_metadata = {
        'name': photo_name,
        'parents': [folder_id]
    }
    media = MediaFileUpload(photo_path, mimetype='image/jpeg')
    drive_service.files().create(body=file_metadata, media_body=media).execute()

def save_emergence_data(emergence_data):
    """Save emergence data to an Excel file."""

    # Convert the emergence data to a DataFrame and save it as an Excel file
    df = pd.DataFrame(emergence_data)
    excel_filename = 'emergence_data.xlsx'
    df.to_excel(excel_filename, index=False)
    print(f"Data saved to {excel_filename}")

    # Upload the emergence data Excel file to Google Drive
    file_metadata = {
        'name': excel_filename,  # The file name with timestamp
        'parents': [parent_folder_file_id]
    }
    media = MediaFileUpload(excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    drive_service.files().create(body=file_metadata, media_body=media).execute()

def is_new_object(x, y, current_time, object_history, history_duration):
    """Check if an object is new based on its coordinates and recent history."""
    for timestamp, hx, hy, hw, hh in object_history:
        # Skip objects that are older than the history duration
        if current_time - timestamp > history_duration:
            continue
        # Check if the object's position is similar to an existing one
        if abs(x - hx) < 50 and abs(y - hy) < 50:
            return False
    return True

def process_frame(frame, bg_subtractor, lower_black, upper_black, object_history, history_duration):
    """Process the frame, detect objects, and update history."""

    # Convert the frame to HSV and apply a color filter
    mask = cv2.inRange(cv2.cvtColor(frame, cv2.COLOR_BGR2HSV), lower_black, upper_black)
    # Apply the background subtractor to detect motion
    motion_mask = cv2.bitwise_and(mask, mask, mask=bg_subtractor.apply(frame))

    # Find contours in the motion mask
    contours, _ = cv2.findContours(motion_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    new_objects = []

    for contour in contours:
        # Filter out contours that are too small or too large
        if not (25 <= cv2.contourArea(contour) < 500):
            continue

        # Get the bounding box of the contour
        x, y, w, h = cv2.boundingRect(contour)
        current_time = datetime.now()

        # Check if the object is new
        if is_new_object(x, y, current_time, object_history, history_duration):
            # Add the new object to the history
            object_history.append((current_time, x, y, w, h))
            # Remove old objects from history
            object_history[:] = [
                obj for obj in object_history if current_time - obj[0] <= history_duration
            ]
            new_objects.append((x, y, w, h))

    return new_objects

def main():

    program_duration_hours = float(input("Enter the program's running duration in hours: "))

    eggs_time_input = input("Enter the time when the Drosophila melanogaster eggs were laid (YYYY-MM-DD HH:MM): ")
    eggs_time = datetime.strptime(eggs_time_input, '%Y-%m-%d %H:%M')

    # Initialize Google Drive
    drive_service, photo_folder_id = initialize_drive()

    # Initialize camera and background subtractor
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: Camera not accessible.")
        return

    # Set up the background subtractor and color thresholds
    bg_subtractor = cv2.createBackgroundSubtractorMOG2(varThreshold=25, detectShadows=False)
    lower_black = np.array([0, 0, 0])
    upper_black = np.array([180, 255, 75])

    # Initialize object history and duration settings
    object_history = []  # Stores (timestamp, x, y, w, h)
    history_duration = timedelta(minutes=5)

    # Convert program duration to seconds
    program_duration = program_duration_hours * 3600
    program_start_time = time.time()
    time_warning_shown = False
    warning_time = program_duration - 60 * 60
    check_photo_interval = 30 * 60  # Interval for check photos in seconds
    next_check_photo_time = program_start_time + check_photo_interval

    # Initialize auto-save timer
    autosave_interval = 10 * 60  # 10 minutes in seconds
    last_autosave_time = time.time()

    print("Program started. Press '1' to extend by 6 hours. Press 'Q' to quit.")

    emergence_data = []  # Store emergence data for upload

    emergence_data.append({
    'time': str(datetime.now().strftime('%Y-%m-%d %H:%M')),
    '': '',  # Placeholder for column 2
    'eggs laid': str(eggs_time)  # This will be written in column 3
})

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        # Flip the frame horizontally for a mirror effect
        frame = cv2.flip(frame, 1)

        # Process the frame and detect new objects
        new_objects = process_frame(frame, bg_subtractor, lower_black, upper_black, object_history, history_duration)
        for x, y, w, h in new_objects:
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M')
            emergence_data.append({'time': current_time})
            print(f"New object detected! Time: {current_time}, Coordinates: ({x}, {y}, {w}, {h})")

            # Draw rectangle around the detected object
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

            # Save photo with the detected object
            photo_name = f'new_object_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.jpg'
            save_photo_to_drive(drive_service, photo_folder_id, frame, photo_name)

        # Save check photos at regular intervals
        if time.time() >= next_check_photo_time:
            check_photo_name = f'check_photo_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.jpg'
            save_photo_to_drive(drive_service, photo_folder_id, frame, check_photo_name)
            next_check_photo_time = time.time() + check_photo_interval

        # Auto-refresh Excel save every 10 minutes
        if time.time() - last_autosave_time >= autosave_interval:
            save_emergence_data(emergence_data)  # Save and overwrite Excel file
            print(f"Autosaved data to Excel at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            last_autosave_time = time.time()

        # Display the current frame
        cv2.imshow('frame', frame)

        elapsed_time = time.time() - program_start_time

        # Check for program duration time
        if elapsed_time >= warning_time and not time_warning_shown:
            print("Warning: Program will close in 1 hour! Enter '1' to extend by 6 hours.")
            time_warning_shown = True

        # Allow extending program duration
        if cv2.waitKey(1) & 0xFF == ord('1'):
            print("Extending program duration by 6 hours.")
            program_duration += 6 * 60 * 60  # Add 6 hours
            warning_time = program_duration - 60 * 60  # Recalculate warning time
            time_warning_shown = False  # Reset flag for the new warning time

        # Exit the program after the set duration or if the user presses 'Q'
        if cv2.waitKey(1) & 0xFF == ord('q') or time.time() - program_start_time >= program_duration:
            # Save and upload the emergence data to Google Drive
            save_emergence_data(emergence_data)
            break

    # Release the camera and close the display window
    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()
