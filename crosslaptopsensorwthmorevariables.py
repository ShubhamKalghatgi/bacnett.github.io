import paho.mqtt.client as mqtt
from paho.mqtt.enums import CallbackAPIVersion
import json
import time

BRAIN_IP = "10.167.64.135" 

client = mqtt.Client(CallbackAPIVersion.VERSION2, "Remote_Sensor_Node")

print(f"Attempting to connect to: {BRAIN_IP}")

try:
    client.connect(BRAIN_IP, 1883, 60)
    print("✅ Connected successfully!")
    
    # Starting conditions
    current_temp = 30.0
    current_humidity = 35.0
    current_co2 = 550
    
    # Target sequence phases: 'COOLING', 'HEATING', 'STABILIZING'
    phase = "COOLING"

    while True:
        # --- LOGIC SEQUENCE ---
        
        if phase == "COOLING":
            status = "Phase 1: Cooling to 18°C"
            current_temp -= 0.5
            current_humidity += 0.5 
            current_co2 += 20
            if current_temp <= 18.0:
                phase = "HEATING"
        
        elif phase == "HEATING":
            status = "Phase 2: Heating to 27°C"
            current_temp += 0.5
            current_humidity -= 0.3
            current_co2 += 20
            if current_temp >= 27.0:
                phase = "STABILIZING"
                
        elif phase == "STABILIZING":
            status = "Phase 3: Stabilizing at 25°C / 1000ppm"
            
            # Adjust Temp toward 25
            if current_temp > 25.1: current_temp -= 0.1
            elif current_temp < 24.9: current_temp += 0.1
            else: current_temp = 25.0
            
            # Adjust CO2 toward 1000
            if current_co2 < 1000: current_co2 += 10
            elif current_co2 > 1010: current_co2 -= 10
            else: current_co2 = 1000
            
            # Keep Humidity in 40-60 range (targeting 50 for stability)
            if current_humidity < 49: current_humidity += 0.5
            elif current_humidity > 51: current_humidity -= 0.5
            else: current_humidity = 50.0

            # --- LOOP RESET TRIGGER ---
            # Check if all parameters have reached the stable target
            if current_temp == 25.0 and current_co2 == 1000 and current_humidity == 50.0:
                print("✨ All parameters stabilized. Restarting cycle in 5 seconds...")
                time.sleep(5) 
                # Reset to starting conditions to loop the test
                current_temp = 30.0
                current_humidity = 35.0
                current_co2 = 550
                phase = "COOLING"

        # --- PAYLOAD CONSTRUCTION ---
        payload = {
            "room": "Studio_1",
            "readings": {
                "co2": int(current_co2),
                "temperature": round(current_temp, 2),
                "humidity": round(current_humidity, 2),
                "hvac_status": status
            }
        }
        
        client.publish("bms/studio_1/sensors", json.dumps(payload))
        print(f"[{phase}] Temp: {payload['readings']['temperature']}°C | CO2: {current_co2} | Hum: {payload['readings']['humidity']}%")
        
        time.sleep(1) # Faster updates for testing the loop
        
except Exception as e:
    print(f"❌ Connection failed: {e}")