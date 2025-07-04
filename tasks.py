class TaskManager:
    def __init__(self):
        self.tasks = {}
        self.next_id = 1

    def add_task(self, name, duration, assigned_to=None):
        task_id = self.next_id
        self.tasks[task_id] = {
            "id": task_id,
            "name": name,
            "duration": duration,
            "assigned_to": assigned_to  # (person, date, shift) or None
        }
        self.next_id += 1
        return task_id

    def delete_task(self, task_id):
        if task_id in self.tasks:
            del self.tasks[task_id]

    def update_task(self, task_id, name=None, duration=None, assigned_to=None):
        if task_id in self.tasks:
            if name is not None:
                self.tasks[task_id]['name'] = name
            if duration is not None:
                self.tasks[task_id]['duration'] = duration
            if assigned_to is not None:
                self.tasks[task_id]['assigned_to'] = assigned_to

    def get_tasks(self):
        return list(self.tasks.values())

    def get_tasks_for(self, person, date, shift):
        result = []
        for task in self.tasks.values():
            assigned = task.get('assigned_to')
            if assigned == (person, date, shift):
                result.append(task)
        return result

    def clear_all_tasks(self):
        self.tasks.clear()
        self.next_id = 1
