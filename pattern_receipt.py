from datetime import datetime

class PatternReceipt:
    def __init__(
        self,
        pattern_name,
        drawing_number,
        impression,
        pattern_type,
        core_box_boolean,
        core_box_name,
        core_box_type,
        shipping_party,
        receiving_party,
        ship_date,
        notes
    ):
        self.pattern_name = pattern_name
        self.drawing_number = drawing_number
        self.impressions = impressions
        self.pattern_type = pattern_type
        self.core_box_boolean = core_box_boolean
        self.core_box_name = core_box_name
        self.core_box_type = core_box_type
        self.shipping_party = shipping_party
        self.drawing_number = drawing_number
        self.ship_date = ship_date
        self.notes = notes

    @classmethod
    def from_user_input(cls, user_input):
        ship_date = datetime.now()
        return cls(
            pattern_name=user_input["pattern_name"],
            drawing_number=user_input["drawing_number"],
            impressions=user_input["impressions"],
            pattern_type=user_input["pattern_type"],
            core_box_boolean=user_input["core_box_boolean"],
            core_box_name=user_input["core_box_name"],
            core_box_type=user_input["core_box_type"],
            shipping_party=user_input["shipping_party"],
            receiving_party=user_input["receiving_party"],
            ship_date=ship_date,
            notes=user_input["notes"]
        )
