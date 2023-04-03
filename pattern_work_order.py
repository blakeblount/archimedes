from datetime import datetime, timedelta

class PatternWorkOrder:
    def __init__(
        self,
        order_number,
        pattern_maker,
        issue_date,
        due_date,
        orderer,
        part_name,
        drawing_number,
        work_requested,
        pattern_type,
        intended_foundry,
        notes,
    ):
        self.order_number = order_number
        self.pattern_maker = pattern_maker
        self.issue_date = issue_date
        self.due_date = due_date
        self.orderer = orderer
        self.part_name = part_name
        self.drawing_number = drawing_number
        self.work_requested = work_requested
        self.pattern_type = pattern_type
        self.intended_foundry = intended_foundry
        self.notes = notes

    @classmethod
    def from_user_input(cls, user_input):
        issue_date = datetime.now()
        due_date = issue_date + timedelta(weeks=12)
        return cls(
            order_number=user_input["order_number"],
            pattern_maker=user_input["pattern_maker"],
            issue_date=issue_date,
            due_date=due_date,
            orderer=user_input["orderer"],
            part_name=user_input["part_name"],
            drawing_number=user_input["drawing_number"],
            work_requested=user_input["work_requested"],
            pattern_type=user_input["pattern_type"],
            intended_foundry=user_input["intended_foundry"],
            notes=user_input["notes"],
        )
