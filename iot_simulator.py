import time
import requests
import random

API_URL = "http://127.0.0.1:5555/api/parking/{id}/status"

def run_iot_simulator():
    print("--------------------------------------------------")
    print("      Smart Parking IoT Sensor Simulator Active    ")
    print("      Simulating Ultrasonic / IR Hardware Nodes   ")
    print("--------------------------------------------------")
    
    slots = [1, 2, 3, 4, 5, 6]
    
    while True:
        # Pick a random slot to trigger an IoT distance sensor change
        target_slot = random.choice(slots)
        
        # Simulate measuring distance (in cm). < 50cm means occupied, > 150cm means vacant/available
        simulated_distance = random.randint(10, 250)
        is_occupied = simulated_distance < 50
        
        status_str = "Occupied" if is_occupied else "Available"
        
        print(f"[IoT Sensor #{target_slot}] Ultrasonic Ping distance: {simulated_distance}cm")
        if is_occupied:
            print(f"[IoT Sensor #{target_slot}] [OCCUPIED] Obstacle detected at {simulated_distance}cm! Bay is OCCUPIED.")
        else:
            print(f"[IoT Sensor #{target_slot}] [FREE] Bay clear. Bay is AVAILABLE.")
            
        # Push state to server via Flask API
        try:
            # We don't specify vehicle number for raw IoT sensors (unlike the AI Camera module)
            payload = {
                "Status": status_str,
                "VehicleNo": "IoT-DETECTED" if is_occupied else None,
                "Force": False # Respect manual admin or user locks
            }
            res = requests.post(API_URL.format(id=target_slot), json=payload, timeout=3)
            data = res.json()
            if "Ignored" in data.get("Message", ""):
                print(f"[IoT Sensor #{target_slot}] Update ignored by server. Slot has a manual booking lock.")
            else:
                print(f"[IoT Sensor #{target_slot}] Server synced successfully! Status: {status_str}")
        except Exception as e:
            print(f"[IoT Sensor Error] Failed to contact server: {e}")
            
        print("Sleeping for 8 seconds before next sensor broadcast...")
        time.sleep(8)

if __name__ == '__main__':
    run_iot_simulator()
