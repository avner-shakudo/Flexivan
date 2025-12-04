import json
from datetime import datetime

# Example movement data for multiple objects
# Each object has a list of (timestamp, lon, lat, height)
data = {
    "ObjectA": [
        ("2024-01-01T00:00:00Z", -74.0, 40.7, 0),
        ("2024-01-01T00:01:00Z", -73.95, 40.71, 0),
        ("2024-01-01T00:02:00Z", -73.9, 40.72, 0),
    ],
    "ObjectB": [
        ("2024-01-01T00:00:00Z", -74.1, 40.69, 0),
        ("2024-01-01T00:01:00Z", -74.05, 40.7, 0),
        ("2024-01-01T00:02:00Z", -74.0, 40.71, 0),
    ]
}

# Colors for objects (RGBA)
colors = {
    "ObjectA": [255, 0, 0, 255],  # red
    "ObjectB": [0, 0, 255, 255],  # blue
}

czml = [{"id": "document", "version": "1.0"}]

for obj_name, movement in data.items():
    start = datetime.fromisoformat(movement[0][0].replace("Z",""))
    
    obj_czml = {
        "id": obj_name,
        "name": obj_name,
        "availability": f"{movement[0][0]}/{movement[-1][0]}",
        "point": {"pixelSize": 12, "color": {"rgba": colors[obj_name]}},
        "path": {
            "material": {
                "polylineOutline": {
                    "color": {"rgba": colors[obj_name]},
                    "outlineColor": {"rgba": [0,0,0,255]},
                    "outlineWidth": 1
                }
            },
            "width": 2,
            "leadTime": 0,
            "trailTime": 3600
        },
        "position": {
            "interpolationAlgorithm": "LINEAR",
            "interpolationDegree": 1,
            "epoch": movement[0][0],
            "cartographicDegrees": []
        }
    }

    for t, lon, lat, h in movement:
        dt = (datetime.fromisoformat(t.replace("Z","")) - start).total_seconds()
        obj_czml["position"]["cartographicDegrees"] += [dt, lon, lat, h]

    czml.append(obj_czml)

# Save CZML
with open("movement_multiple.czml", "w") as f:
    json.dump(czml, f, indent=2)

print("✔ CZML file created: movement_multiple.czml")
