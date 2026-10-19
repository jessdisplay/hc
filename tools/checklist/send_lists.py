"""Who each form goes out to, and what they have to do with it.

Read from each form's own 'Completed by' and 'Signed by' (the 16 Oct audit of
Harmony's documents). A form appears once for every person it goes to: a
bowel care support plan is written by the health practitioner and signed by
the participant, so it sits on both lists. Forms on no list are the ones
Harmony fills in itself; make_workbook.py lists those on their own tab.

Each entry is (document title, what they do). Every title must match a
document exactly; make_workbook.py stops if one does not.
"""

FILL_SIGN, FILL, SIGN = "Fill in and sign", "Fill in", "Sign"
READ_SIGN, READ_ACCEPT = "Read and sign", "Read, then accept it in the Worker Declarations"

PARTICIPANT = {
    "tab": "Send to participants",
    "colour": "Participant",
    "title": "Send to the participant, or their guardian or representative",
    "intro": "What a participant fills in or signs. A Harmony staff member explains each form and countersigns "
             "where the form asks for it.",
    "groups": [
        ("Before supports start", [
            ("Client Intake Form", FILL_SIGN),
            ("Privacy Consent Form", FILL_SIGN),
            ("Services Agreement", SIGN),
            ("Services and Supports Schedule", SIGN),
            ("Client Support Plan", SIGN),
            ("Risk Assessment Form", SIGN),
        ]),
        ("Only if it applies to them", [
            ("Advocacy or Support Person Request Form", FILL_SIGN),
            ("Money and Property Consent Form", FILL_SIGN),
            ("Risk Taking Form", SIGN),
            ("Mealtime Management Plan", SIGN),
            ("Complex Bowel Care Support Plan", SIGN),
            ("Enteral Feeding Support Plan", SIGN),
            ("Urinary Catheter Support Plan", SIGN),
            ("Complex Wound Management Support Plan", SIGN),
            ("Support Coordination Services Agreement", SIGN),
            ("Conflict of Interest Declaration (support coordination)", SIGN),
        ]),
        ("Supported independent living (SIL)", [
            ("Services Agreement (SIL)", SIGN),
            ("Home and Support Arrangements Schedule to the Services Agreement (SIL)", SIGN),
            ("SDA and SIL Collaboration Agreement", SIGN),
            ("SIL Participant Support Plan", SIGN),
            ("SIL Participant Risk Assessment and Safeguarding Review", SIGN),
            ("SIL Supported Decision-Making Assessment and Record", SIGN),
            ("SIL Conflict of Interest Declaration and Register", SIGN),
        ]),
        ("Later on", [
            ("Change to Supports Agreement", SIGN),
            ("Client Survey", FILL),
            ("Feedback and Complaints Form", FILL),
            ("Exit and Transition Form", SIGN),
        ]),
    ],
}

WORKER = {
    "tab": "Send to workers",
    "colour": "Worker",
    "title": "Send to each worker",
    "intro": "What a worker fills in, signs or reads. Their supervisor or the Principal completes the rest of "
             "each form and signs it with them.",
    "groups": [
        ("Before they start", [
            ("Pre Employment Collection Form", FILL_SIGN),
            ("Worker Declarations", FILL_SIGN),
            ("Position Description: Disability Support Worker", READ_ACCEPT),
            ("Position Description: Disability Support Worker (second version)", READ_ACCEPT),
            ("Position Description: Disability Support Worker (Employment Consultant)", READ_ACCEPT),
            ("Position Description: Disability Support Worker (Travel and Transport)", READ_ACCEPT),
            ("Position Description: Disability Support Worker (Implementing Behaviour Supports)", READ_ACCEPT),
            ("Position Description: Support Coordinator", READ_ACCEPT),
            ("Staff Induction Checklist", SIGN),
        ]),
        ("Training and supervision", [
            ("Training and Development Register", "Sign each training line"),
            ("Training Feedback Form", FILL),
            ("Worker Supervision Record", SIGN),
            ("HR Performance Appraisal", "Fill in the self-evaluation, then sign"),
            ("Staff Performance Improvement Plan", "Sign (only if a plan is needed)"),
            ("High Intensity Worker Training Plan and Competency Record", SIGN),
            ("Training and Evaluation Form (behaviour support knowledge test)", FILL),
            ("Coaching Form for Positive Behaviour Support Plan", SIGN),
            ("Training and Development Register (Implementing Behaviour Supports)", "Sign each training line"),
        ]),
        ("Working in a SIL house", [
            ("SIL Worker Training Plan", SIGN),
            ("SIL Worker Competency Assessment Record", SIGN),
            ("SIL House-Level Safeguarding Plan", READ_SIGN),
            ("SIL House Emergency Plan", READ_SIGN),
            ("SIL House Vision, Values and Service Delivery Statement", READ_SIGN),
            ("SIL Supported Decision-Making Record", "Fill in an entry each time, and sign the declaration"),
            ("SIL Conflict of Interest Declaration and Register", "Fill in and sign (only if they have an interest)"),
        ]),
        ("On every shift (kept in the participant's file)", [
            ("Worker Timesheets (NDIS Support Record)", FILL_SIGN),
            ("Progress Notes Template", FILL),
            ("Administering Medication Checklist and Medication Chart", "Initial each dose given"),
            ("Collection and Storage of Medicine Register", "Sign each entry"),
            ("Client Cash Reconciliation", "Fill in (when they handle the participant's money)"),
            ("Daily Fluid Balance Chart", "Fill in (initial each entry)"),
            ("Bowel Monitoring Chart", FILL),
            ("Enteral Feeding and Tube Flush Chart", FILL),
            ("Urinary Catheter Documentation Chart", FILL_SIGN),
            ("Wound Observation and Dressing Chart", "Fill in (initial each entry)"),
            ("ABC Data Form", "Fill in (each time a behaviour of concern happens)"),
            ("Waste Management Report", FILL_SIGN),
            ("Waste Removal Records Register", FILL),
        ]),
        ("When they leave", [
            ("Staff Exit Checklist", SIGN),
        ]),
    ],
}

PRACTITIONER = {
    "tab": "Send to practitioners",
    "colour": "Organisation",
    "title": "Send to health practitioners, behaviour support practitioners and partners",
    "intro": "People outside Harmony who write, fill in or sign part of a participant's paperwork.",
    "groups": [
        ("Qualified health practitioner (nurse, speech pathologist, dietitian)", [
            ("Complex Bowel Care Support Plan", "Write the plan with the participant, and sign"),
            ("Enteral Feeding Support Plan", "Write the plan with the participant, and sign"),
            ("Urinary Catheter Support Plan", "Write the plan with the participant, and sign"),
            ("Complex Wound Management Support Plan", "Write the plan with the participant, and sign"),
            ("Mealtime Management Plan", "Write the plan, and sign"),
            ("High Intensity Worker Training Plan and Competency Record", "Train and assess the worker, and sign"),
        ]),
        ("Prescriber (the participant's doctor)", [
            ("Services Medication and Participant Details", "Fill in the prescriber's details, and sign"),
            ("Administering Medication Checklist and Medication Chart", "Write and sign each medication order"),
        ]),
        ("Behaviour support practitioner", [
            ("Template Letter to PBSP (behaviour support practitioner)", "Harmony sends this letter to them"),
            ("Behaviour Support Plan (Comprehensive)", "Write the plan, and sign"),
            ("Behaviour Support Plan (Interim)", "Write the plan, and sign"),
            ("Coaching Form for Positive Behaviour Support Plan", "Fill in section A"),
            ("Positive Behaviour Support Plan Training Report", "Fill in and sign after training the workers"),
        ]),
        ("SDA provider (the participant's housing provider)", [
            ("SDA and SIL Collaboration Agreement", SIGN),
        ]),
    ],
}

SEND_TABS = [PARTICIPANT, WORKER, PRACTITIONER]
