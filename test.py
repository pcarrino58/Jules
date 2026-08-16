def GetBestJobPlanMatch(aType, aSub, jNorms):
    aType = aType.upper().strip()
    aSub = aSub.upper().strip()

    aType = aType.replace("SWITCH GEAR", "SWITCHGEAR").replace("TANKS", "TANK")
    aType = aType.replace("PUMPS", "PUMP").replace("VALVES", "VALVE")
    aSub = aSub.replace("SWITCH GEAR", "SWITCHGEAR").replace("TANKS", "TANK")
    aSub = aSub.replace("PUMPS", "PUMP").replace("VALVES", "VALVE")

    # Ignore dimensions
    if aSub.startswith("SIZE") or "X" in aSub or "LB" in aSub:
        aSub = ""

    aFull = (aType + " " + aSub).strip()
    aFullRev = (aSub + " " + aType).strip()

    for key in jNorms:
        jNorms[key] = key.upper().strip()
        jNorms[key] = jNorms[key].replace("SWITCH GEAR", "SWITCHGEAR").replace("TANKS", "TANK")
        jNorms[key] = jNorms[key].replace("PUMPS", "PUMP").replace("VALVES", "VALVE")
        jNorms[key] = jNorms[key].replace("DAMPERS", "DAMPER").replace("WATER HEATERS", "WATER HEATER")
        jNorms[key] = jNorms[key].replace("AIR TERMINALS", "AIR TERMINAL")

    # Priority 1
    if aFull != "":
        for key in jNorms:
            if aFull == jNorms[key]: return key

    # Priority 2
    if aFullRev != "":
        for key in jNorms:
            if aFullRev == jNorms[key]: return key

    # Priority 3
    if aType != "":
        for key in jNorms:
            if aType == jNorms[key]: return key

    # Priority 4
    if aSub != "":
        for key in jNorms:
            if aSub == jNorms[key]: return key

    # Step 2
    for key in jNorms:
        jpDesc = key
        jNorm = jNorms[key]

        if (aType == "VAV" or aSub == "VAV" or aType == "VAV BOX") and jNorm == "AIR TERMINAL":
            return jpDesc
        elif aType == "PUMP" and "CENTRIFUGAL" in aSub and jNorm == "PUMP INSPECTION":
            return jpDesc
        elif aType == "PUMP" and aSub == "" and jNorm == "PUMP INSPECTION":
            return jpDesc
        elif aType == "SUMP PUMP" and jNorm == "SUMP PUMP INSPECTION":
            return jpDesc
        elif aType == "TANK" and "EXPANSION" in aSub and jNorm == "TANK - EXPANSION CUSHION":
            return jpDesc
        elif aType == "TANK" and "CONDENSATE" in aSub and jNorm == "TANK - CONDENSATE":
            return jpDesc
        elif aType == "FUEL TANK" and (jNorm == "DIESEL DAY TANK INSPECTION" or jNorm == "DIESEL DAY TANK"):
            return jpDesc
        elif aType == "SWITCHGEAR" and (jNorm == "SWITCHGEAR" or jNorm == "SWITCHGEAR, HIGH VOLTAGE"):
            return jpDesc
        elif aType == "TRANSFORMER" and jNorm == "TRANSFORMER INSPECTION":
            return jpDesc
        elif (aType == "WATER HEATER" or aType == "HOT WATER HEATER") and jNorm == "WATER HEATER":
            return jpDesc
        elif aType == "VFD" and jNorm == "VFD INSPECTION":
            return jpDesc
        elif aType == "SPRINKLER" and jNorm == "SPRINKLER SYSTEM TESTING":
            return jpDesc
        elif aType == "CARD READER" and jNorm == "SECURITY SYSTEMS":
            return jpDesc
        elif aType == "MOTOR CONTROL CENTER" and jNorm == "MCC":
            return jpDesc
        elif aType == "DAMPER" and jNorm == "DAMPER":
            return jpDesc
        elif (aType == "BACK FLOW PREVENTION DEVICE" or aType == "BACKFLOW PREVENTION DEVICE" or aFull == "BACK FLOW PREVENTION DEVICE") and jNorm == "BACKFLOW PREVENTER":
            return jpDesc
        elif (aType.startswith("ELEVATOR TRACTION") or aFull.startswith("ELEVATOR TRACTION")) and jNorm == "TRACTION ELEVATOR PM TASKS":
            return jpDesc
        elif aType == "FIRE EXTINGUISHER" and "EXTINGUISHER" in jNorm:
            return jpDesc

    # Step 3
    if aFull != "" and " " in aFull:
        for key in jNorms:
            if aFull in jNorms[key]: return key

    if aType != "" and aSub != "":
        for key in jNorms:
            if aType in jNorms[key] and aSub in jNorms[key]: return key

    if aType != "" and " " in aType:
        for key in jNorms:
            if aType in jNorms[key]: return key

    if aSub != "" and " " in aSub:
        for key in jNorms:
            if aSub in jNorms[key]: return key

    return ""

jNorms = {
    "Air Conditioners - Packaged Split AC System": "",
    "Air Conditioners - Window unit": "",
    "By-Pass Filter": "",
    "Air Compressor": "",
    "Heaters - Unit heater hot water": "",
    "Domestic Hot Water Heater": "",
    "Fire Extinguisher - Inspection": ""
}

print("Match:", GetBestJobPlanMatch("Fire Extinguisher", "CO2 BC 15 lb", jNorms))
