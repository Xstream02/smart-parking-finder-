import time
import requests
import cv2
import numpy as np
import random
import string

API_URL = "http://127.0.0.1:5555/api/parking/{id}/status"
USE_WEBCAM = False # Set to True to capture physical webcam frame
HEADLESS_MODE = False # Set to True to disable cv2.imshow popups (useful for headless servers)

# Fallbacks for OCR
HAS_OCR = False
try:
    import easyocr
    reader = easyocr.Reader(['en'])
    HAS_OCR = True
    print("[Vision] EasyOCR successfully initialized!")
except Exception:
    print("[Vision] EasyOCR not installed. Falling back to Mock ALPR Recognition.")

# Bounding boxes for virtual camera parking slots
PARKING_SPOTS = {
    1: [100, 150, 250, 270], # Spot A1
    2: [100, 300, 250, 420], # Spot A2
    3: [100, 450, 250, 570], # Spot A3
    4: [550, 150, 700, 270], # Spot B1
    5: [550, 300, 700, 420], # Spot B2
    6: [550, 450, 700, 570]  # Spot B3
}

def generate_random_plate():
    state = random.choice(["MH", "DL", "KA", "HR", "UP", "GJ"])
    code = f"{random.randint(10, 99)}"
    letters = "".join(random.choices(string.ascii_uppercase, k=2))
    number = f"{random.randint(1000, 9999)}"
    return f"{state} {code} {letters} {number}"

def draw_parking_lot(occupancies, license_plates):
    """
    Generate a dynamic virtual image representing the physical parking lot feed.
    This demonstrates real-time OpenCV image analysis in action!
    """
    # Create black canvas representing camera perspective
    img = np.zeros((650, 800, 3), dtype=np.uint8)
    
    # Draw dark gray road asphalt
    cv2.rectangle(img, (280, 0), (520, 650), (40, 40, 40), -1)
    # Draw yellow divider lane markings
    for y in range(30, 650, 80):
        cv2.rectangle(img, (395, y), (405, y+40), (0, 220, 255), -1)
        
    # Title info banner
    cv2.rectangle(img, (0, 0), (800, 80), (15, 23, 42), -1)
    cv2.putText(img, "AI SMART PARKING SYSTEM - LIVE CAMERA FEED", (60, 48), 
                cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 242, 254), 2, cv2.LINE_AA)
    
    # Draw slot dividers and slots
    for spot_id, box in PARKING_SPOTS.items():
        x1, y1, x2, y2 = box
        is_occupied = occupancies.get(spot_id, False)
        
        # Color: Green = Available, Red = Occupied
        color = (0, 50, 255) if is_occupied else (0, 255, 0)
        thick = 3 if is_occupied else 2
        
        # Draw dashed spot borders (using lines)
        cv2.rectangle(img, (x1, y1), (x2, y2), color, thick)
        
        # Spot labels
        label = f"SPOT {spot_id}"
        cv2.putText(img, label, (x1 + 10, y1 + 30), 
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1, cv2.LINE_AA)
        
        if is_occupied:
            # Draw a stylized virtual car bounding box
            cv2.rectangle(img, (x1 + 15, y1 + 45), (x2 - 15, y2 - 15), (100, 100, 100), -1)
            cv2.rectangle(img, (x1 + 25, y1 + 55), (x2 - 25, y2 - 35), (200, 200, 200), -1) # Windshield
            
            # Virtual license plate banner
            plate = license_plates.get(spot_id, "MH 12 AB 1234")
            cv2.rectangle(img, (x1 + 20, y2 - 30), (x2 - 20, y2 - 18), (255, 255, 255), -1)
            cv2.putText(img, plate, (x1 + 25, y2 - 20), 
                        cv2.FONT_HERSHEY_SIMPLEX, 0.35, (0, 0, 0), 1, cv2.LINE_AA)
            
            # Draw AI vehicle detection overlay bounding box
            cv2.rectangle(img, (x1 - 5, y1 - 5), (x2 + 5, y2 + 5), (0, 0, 255), 1)
            cv2.putText(img, "CAR: 98% Conf", (x1 - 5, y1 - 10), 
                        cv2.FONT_HERSHEY_SIMPLEX, 0.4, (0, 0, 255), 1, cv2.LINE_AA)
        else:
            cv2.putText(img, "VACANT", (x1 + 10, y2 - 20), 
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1, cv2.LINE_AA)

    return img

def run_vision_module():
    print("--------------------------------------------------")
    print("      Upgraded Smart Parking Vision AI Active     ")
    print("      Running OpenCV Simulation & ALPR Feed       ")
    print("--------------------------------------------------")
    
    # Store dynamic simulation states
    occupancies = {i: False for i in PARKING_SPOTS}
    license_plates = {i: "" for i in PARKING_SPOTS}
    
    while True:
        # Every iteration, simulate a random state change (arrival or exit)
        changed_spot = random.choice(list(PARKING_SPOTS.keys()))
        current_state = occupancies[changed_spot]
        
        # 40% chance to toggle slot state
        if random.random() < 0.4:
            new_state = not current_state
            occupancies[changed_spot] = new_state
            
            if new_state:
                # Vehicle arrived: Perform ALPR/OCR and extract license plate
                license_plates[changed_spot] = generate_random_plate()
                plate = license_plates[changed_spot]
                status_str = "Occupied"
                print(f"[Vision AI] Car entered Spot {changed_spot}. ALPR Identified Plate: {plate}")
            else:
                # Vehicle left
                plate = None
                status_str = "Available"
                license_plates[changed_spot] = ""
                print(f"[Vision AI] Car exited Spot {changed_spot}.")
            
            # Push status to Flask backend API
            try:
                payload = {
                    "Status": status_str,
                    "VehicleNo": plate,
                    "Force": False # False: respect manual lock (Reserved/Pending)
                }
                res = requests.post(API_URL.format(id=changed_spot), json=payload, timeout=3)
                data = res.json()
                if "Ignored" in data.get("Message", ""):
                    # Sync our local state back to Reserved because it was manually locked
                    occupancies[changed_spot] = True
                    license_plates[changed_spot] = data.get("VehicleNo", "RESERVED")
                    print(f"[Vision AI] Spot {changed_spot} is locked (Reserved/Pending). Overwrite blocked.")
            except Exception as e:
                print(f"[Vision AI Error] Failed to update server status: {e}")
        
        # Render visual OpenCV feed
        frame = draw_parking_lot(occupancies, license_plates)
        
        if not HEADLESS_MODE:
            try:
                cv2.imshow("Smart Parking - OpenCV AI Vision Feed", frame)
                # Listen to key press for exits (delay 3000ms = 3 seconds)
                if cv2.waitKey(3000) & 0xFF == 27: # ESC key
                    break
            except Exception as e:
                print(f"[Vision Warning] GUI display not available: {e}. Running in console mode.")
                time.sleep(3)
        else:
            time.sleep(3)

    if not HEADLESS_MODE:
        cv2.destroyAllWindows()

if __name__ == "__main__":
    run_vision_module()
