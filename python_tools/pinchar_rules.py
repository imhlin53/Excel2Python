"""
Rules for each Pin Type.
"""

PINCHAR_RULES = {

    "Motor": {
        "formula_rules":{
            "i3":{
                "source": "i2",
                "multiplier": 4
            }
        }
    },
    
    "Kapazität": {
        "formula_rules": {
            "i3":{
                "source": "i2",
                "multiplier": 5
            }
        }
    },
    
    "Glühlampe": {
        "formula_rules":{
            "i3": {
                "source": "i2",
                "multiplier": 3.5
            }
        }
    },
    
    "spule": {
        "formula_rules": {
            "i2": {
                "source": "i1",
                "multiplier": 4
            }
        }
    },
    
    "Nicht verbunden": {

        "required_fields": [],

        "fixed_values": {
            "i1": 0
        }
    },

    "Senke (SG Last)": {

        "required_fields": [
            "i1",
            "i2",
            "i3"
        ]
    },

    "Signal (CAN, Messein/-ausgang)": {

        "required_fields": [
            "i1",
            "i2"
        ],
        
        "minimum_values": {
            "i1": 0,
            "i2": 0
        }
    },

    "Senke (Komponente)": {

        "required_fields": [
            "i1",
            "i2"
        ]
    }
}