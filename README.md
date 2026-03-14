## GENAI TEST
Name: Aero_Hub_Plate
Prompt: Create a circular base with radius P1 and thickness P3. Add a central cylindrical boss on top with radius P2 and height P3*2. Subtract a hole through the very center with diameter P4 that goes through the entire part. Finally, subtract four small bolt holes with radius 5 at the cardinal points (N, S, E, W) of the base plate, positioned at a distance of P1*0.75 from the center.

## Name: Sensor
Prompt: I need a rectangular base block where the length is P1, width is P2, and height is P3. I need a circular pocket cut into the top center to hold the sensor. The pocket diameter is P4, and it should only be drilled halfway down into the block, not all the way through.

Name: Radial_Cooling_Hub
Prompt: Create a central cylindrical core with radius P2 and height P3. Subtract a central bore hole with diameter P4 through the core. Now, create a radial array of cooling fins: add 8 thin rectangular boxes, each with length P1, width 5, and height P3. Position each box so it radiates from the center. Space them at 45-degree intervals. Finally, subtract a sphere from the very top center with radius P2 to create a smooth aerodynamic dish effect on the top of the hub.

Name:Bio_Lightweight_Bracket
Prompt:Create a complex, bio-inspired bracket. Start with a robust L-shaped bracket base: a vertical plate (length P1, width 10, height P3) joined to a horizontal plate (length P1, width P2, height 10). Now, perform 'material reduction' for a lightweight, optimized design. Subtract an array of 5 large, hexagonal-shaped cylinder holes through the face of the vertical plate. Position them in a zigzag pattern to maximize structural integrity while minimizing weight. Then, subtract a series of 3 smaller, organic, pocket-style sphere holes from the horizontal base plate. Finally, subtract a precise mounting through-hole (diameter P4) at each end of the vertical plate.

Name: Aerospace_Gear_Blank
Prompt: Create a high-performance gear hub. Start with a cylindrical base (radius P1, height P3) representing the outer rim. Now, skeletonize it for weight reduction: Subtract a cylindrical ring between radius P4 and P1*0.8, leaving only a outer rim and a central hub. Now, add 6 structural 'spokes' connecting the center to the rim: each spoke is a box (length P1, width 8, height P3) rotated at 60-degree intervals. Finally, subtract a precision keyed-bore from the center: a cylinder hole (radius 15, height P3+10) with a small rectangular keyway box (width 5, length 5, height P3+10) subtracted at the top of the bore.

Name: Serrated_Lock_Washer
Prompt: Create a flat ring with outer radius P1, inner radius P4, and thickness P3. Now, create a serrated edge: Subtract 36 small boxes around the outer perimeter. Each box should be width 5, length 10, and height P3. Rotate each box by 10 degrees relative to the previous one to create a saw-tooth pattern.



ENTERPRISE_TEMPLATES = {
    "Bearing_Housing": {
        "description": "Standard industrial bearing housing. P1=Outer OD, P2=Bore ID, P3=Length.",
        "operations": [
            {"type": "add", "shape": "cylinder", "params": {"r": "P1/2", "h": "P3", "x": "0", "y": "0", "z": "0"}},
            {"type": "subtract", "shape": "cylinder", "params": {"r": "P2/2", "h": "P3", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "Drone_Motor_Mount": {
        "description": "Quad-copter motor mounting bracket. P1=Base Width, P2=Thickness, P3=Center Hole Dia.",
        "operations": [
            {"type": "add", "shape": "box", "params": {"l": "P1", "w": "P1", "h": "P2", "x": "0", "y": "0", "z": "0"}},
            {"type": "subtract", "shape": "cylinder", "params": {"r": "P3/2", "h": "P2", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "Aerospace_Flange": {
        "description": "High-pressure pipe flange. P1=Flange OD, P2=Pipe ID, P3=Thickness.",
        "operations": [
            {"type": "add", "shape": "cylinder", "params": {"r": "P1/2", "h": "P3", "x": "0", "y": "0", "z": "0"}},
            {"type": "subtract", "shape": "cylinder", "params": {"r": "P2/2", "h": "P3", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "Robotics_Chassis_Plate": {
        "description": "Main base plate for an AGV robot. P1=Length, P2=Width, P3=Thickness.",
        "operations": [
            {"type": "add", "shape": "box", "params": {"l": "P1", "w": "P2", "h": "P3", "x": "0", "y": "0", "z": "0"}},
            # Hollow out the center slightly for weight reduction
            {"type": "subtract", "shape": "box", "params": {"l": "P1*0.6", "w": "P2*0.6", "h": "P3", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "Splined_Shaft": {
        "description": "Power transmission shaft. P1=Length, P2=Diameter, P3=Spline Depth.",
        "operations": [
            {"type": "add", "shape": "cylinder", "params": {"r": "P2/2", "h": "P1", "x": "0", "y": "0", "z": "0"}},
            # Simulating a spline groove
            {"type": "subtract", "shape": "box", "params": {"l": "P2+2", "w": "P3", "h": "P1", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "Sensor_Bracket_L": {
        "description": "L-Bracket for optical sensors. P1=Height, P2=Base Length, P3=Thickness.",
        "operations": [
            {"type": "add", "shape": "box", "params": {"l": "P3", "w": "P2", "h": "P1", "x": "-P3/2", "y": "P2/2", "z": "0"}},
            {"type": "add", "shape": "box", "params": {"l": "P2", "w": "P3", "h": "P3", "x": "P2/2", "y": "0", "z": "0"}}
        ]
    },
    "Pulley_Wheel": {
        "description": "V-belt pulley wheel. P1=Outer Dia, P2=Thickness, P3=Bore Dia.",
        "operations": [
            {"type": "add", "shape": "cylinder", "params": {"r": "P1/2", "h": "P2", "x": "0", "y": "0", "z": "0"}},
            {"type": "subtract", "shape": "cylinder", "params": {"r": "P3/2", "h": "P2", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "Heat_Sink_Base": {
        "description": "Thermal dissipation block. P1=Length, P2=Width, P3=Fin Height.",
        "operations": [
            {"type": "add", "shape": "box", "params": {"l": "P1", "w": "P2", "h": "P3", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "Hydraulic_Piston_Cap": {
        "description": "End cap for hydraulic cylinders. P1=Outer Dia, P2=Inner Dia, P3=Height.",
        "operations": [
            {"type": "add", "shape": "cylinder", "params": {"r": "P1/2", "h": "P3", "x": "0", "y": "0", "z": "0"}},
            {"type": "subtract", "shape": "cylinder", "params": {"r": "P2/2", "h": "P3/2", "x": "0", "y": "0", "z": "P3/2"}}
        ]
    },
    "Motor_Coupling": {
        "description": "Shaft coupling for NEMA motors. P1=Outer Dia, P2=Length, P3=Bore Dia.",
        "operations": [
            {"type": "add", "shape": "cylinder", "params": {"r": "P1/2", "h": "P2", "x": "0", "y": "0", "z": "0"}},
            {"type": "subtract", "shape": "cylinder", "params": {"r": "P3/2", "h": "P2", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "Conveyor_Roller": {
        "description": "Material handling roller. P1=Roller Dia, P2=Length, P3=Axle Dia.",
        "operations": [
            {"type": "add", "shape": "cylinder", "params": {"r": "P1/2", "h": "P2", "x": "0", "y": "0", "z": "0"}},
            {"type": "subtract", "shape": "cylinder", "params": {"r": "P3/2", "h": "P2", "x": "0", "y": "0", "z": "0"}}
        ]
    },
    "CNC_Fixture_Plate": {
        "description": "Tooling plate for 5-axis machines. P1=Length, P2=Width, P3=Thickness.",
        "operations": [
            {"type": "add", "shape": "box", "params": {"l": "P1", "w": "P2", "h": "P3", "x": "0", "y": "0", "z": "0"}}
        ]
    }
}
