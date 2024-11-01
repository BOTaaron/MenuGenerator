class ExclusionManager:
    def __init__(self):
        self.default_exclusions = ["Misc drink - £0.00"]
        self.exclusions = set(self.default_exclusions)

    def apply_exclusions(self, departments):
        filtered_departments = {}

        for dept, items in departments.items():
            filtered_items = [item for item in items if item['name'] not in self.exclusions]
            if filtered_items:
                filtered_departments[dept] = filtered_items

        return filtered_departments
