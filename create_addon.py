import re
from typing import List, Dict

def create_addon(text: str) -> List[Dict[str, str]]:
    data = {
    # Addon Covers - Binary options (Yes/No)
    "Ambulance Cover": "",
    "Anyone Illness": "",
    "Attendant Care": "",
    "Cancer Cover": "",
    "Convalescence Benefit": "",
    "Critical Illness Benefit": "",
    "Daily/Hospital Cash Benefit": "",
    "Dental Cover": "",
    "Diabetic Cover": "",
    "Doctor & Nurse Home Visit Cover": "",
    "Education Fund": "",
    "Funeral": "",
    "Getwell Benefit": "",
    # Additional Addon Covers from new image
    "Hardship Critical Illness Cover": "",
    "Health Check up": "",
    "Hypertension Cover": "",
    "Intensive Care Benefit": "",
    "Loss Of Pay Cover": "",
    "Medical Evacuation Cover": "",
    "Medical Second Opinion": "",
    "Non Medical Expense Cover": "",
    "Out Patient Cover": "",
    "Optical Cover": "",
    "Organ Donor Medical Expense Cover": "",
    # Additional Addon Covers from third image
    "Personal Accident Cover": "",
    "Pre Existing Disease Benefit": "",
    "Psychiatric Cover": "",
    "Recovery Benefit": "",
    "Referral Hospital Care": "",
    "Surgical Benefit": "",
    "Top Up Cover": "",
    "Vaccination/Immunization Cover": ""
    }

    # Check for Ambulance Cover
    if "ambulance" in text.lower() or "emergency ambulance" in text.lower():
        data["Ambulance Cover"] = "Yes"
        

    # Check for Anyone Illness Benefit
    if "anyone illness" in text.lower() or "any illness" in text.lower():
        data["Anyone Illness"] = "Yes"
        

    # Check for Attendant Care
    if "attendant" in text.lower() or "attendance" in text.lower():
        data["Attendant Care"] = "Yes"
        

    # Check for Cancer Cover
    if "cancer" in text.lower() or "oncology" in text.lower():
        data["Cancer Cover"] = "Yes"
        

    # Check for Convalescence Benefit
    if "convalescence" in text.lower() or "convalescent" in text.lower():
        data["Convalescence Benefit"] = "Yes"
        

    # Check for Critical Illness Benefit
    if "critical illness" in text.lower() or "critical disease" in text.lower():
        data["Critical Illness Benefit"] = "Yes"
        

    # Check for Daily/Hospital Cash Benefit
    if "daily cash" in text.lower() or "hospital cash" in text.lower() or "cash benefit" in text.lower():
        data["Daily/Hospital Cash Benefit"] = "Yes"
        

    # Check for Dental Cover
    if "dental" in text.lower() or "dental treatment" in text.lower():
        data["Dental Cover"] = "Yes"
        

    # Check for Diabetic Cover
    if "diabetic" in text.lower() or "diabetes" in text.lower():
        data["Diabetic Cover"] = "Yes"
        

    # Check for Doctor & Nurse Home Visit Cover
    if "home visit" in text.lower() or "doctor visit" in text.lower() or "nurse visit" in text.lower():
        data["Doctor & Nurse Home Visit Cover"] = "Yes"


    # Check for Education Fund
    if "education" in text.lower() or "education fund" in text.lower():
        data["Education Fund"] = "Yes"
        

    # Check for Funeral
    if "funeral" in text.lower() or "funeral expenses" in text.lower():
        data["Funeral"] = "Yes"
        

    # Check for Getwell Benefit
    if "getwell" in text.lower() or "get well" in text.lower():
        data["Getwell Benefit"] = "Yes"
        

    # Check for Hardship Critical Illness


    # Check for Hardship Critical Illness Cover (additional)
    if "hardship critical illness cover" in text.lower():
        data["Hardship Critical Illness Cover"] = "Yes"
        

    # Check for Health Check up
    if "health check" in text.lower() or "health checkup" in text.lower() or "health screening" in text.lower():
        data["Health Check up"] = "Yes"
        

    # Check for Hypertension Cover
    if "hypertension" in text.lower() or "high blood pressure" in text.lower():
        data["Hypertension Cover"] = "Yes"
        

    # Check for Intensive Care Benefit
    if "intensive care" in text.lower() or "icu" in text.lower():
        data["Intensive Care Benefit"] = "Yes"
        

    # Check for Loss Of Pay Cover
    if "loss of pay" in text.lower() or "loss of income" in text.lower():
        data["Loss Of Pay Cover"] = "Yes"
        

    # Check for Medical Evacuation Cover
    if "medical evacuation" in text.lower() or "evacuation" in text.lower():
        data["Medical Evacuation Cover"] = "Yes"
        

    # Check for Medical Second Opinion
    if "second opinion" in text.lower() or "medical opinion" in text.lower():
        data["Medical Second Opinion"] = "Yes"
        

    # Check for Non Medical Expense Cover
    if "non medical" in text.lower() or "non-medical" in text.lower():
        data["Non Medical Expense Cover"] = "Yes"
        

    # Check for Out Patient Cover
    if "out patient" in text.lower() or "outpatient" in text.lower() or "opd" in text.lower():
        data["Out Patient Cover"] = "Yes"
       

    # Check for Optical Cover
    if "optical" in text.lower() or "eye care" in text.lower() or "vision" in text.lower():
        data["Optical Cover"] = "Yes"
        

    # Check for Organ Donor Medical Expense Cover
    if "organ donor" in text.lower() or "organ donation" in text.lower():
        data["Organ Donor Medical Expense Cover"] = "Yes"
       

    # Check for Personal Accident Cover
    if "personal accident" in text.lower() or "accident cover" in text.lower():
        data["Personal Accident Cover"] = "Yes"
     

    # Check for Pre Existing Disease Benefit
    if "pre existing" in text.lower() or "pre-existing" in text.lower() or "existing disease" in text.lower():
        data["Pre Existing Disease Benefit"] = "Yes"
     

    # Check for Psychiatric Cover
    if "psychiatric" in text.lower() or "psychiatry" in text.lower() or "mental health" in text.lower():
        data["Psychiatric Cover"] = "Yes"
    

    # Check for Recovery Benefit
    if "recovery benefit" in text.lower() or "recovery" in text.lower():
        data["Recovery Benefit"] = "Yes"
       

    # Check for Referral Hospital Care
    if "referral hospital" in text.lower() or "hospital referral" in text.lower():
        data["Referral Hospital Care"] = "Yes"
      

    # Check for Surgical Benefit
    if "surgical benefit" in text.lower() or "surgery" in text.lower():
        data["Surgical Benefit"] = "Yes"
        

    # Check for Top Up Cover
    if "top up" in text.lower() or "topup" in text.lower():
        data["Top Up Cover"] = "Yes"
      

    # Check for Vaccination/Immunization Cover
    if "vaccination" in text.lower() or "immunization" in text.lower() or "vaccine" in text.lower():
        data["Vaccination/Immunization Cover"] = "Yes"
       
    # Set default "No" for addons not found
    for key in data:
        if data[key] == "":
            data[key] = "No"

    endorsement_patterns = re.findall(r"Endt\. No\. \d+[a-z]?", text)
    if endorsement_patterns:
        print(f"\nEndorsements found: {', '.join(set(endorsement_patterns))}")
    
    special_conditions = re.findall(r"Special Condition[s]?.*?[A-Z]", text, re.IGNORECASE)
    if special_conditions:
        print(f"\nSpecial conditions found: {', '.join(set(special_conditions))}")



    return [data]