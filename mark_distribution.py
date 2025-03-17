import pandas as pd
import numpy as np
from pulp import *
from typing import Dict, List, Tuple

class MarkDistributor:
    def __init__(self):
        self.grade_points = {
            'F': 0, 'E': 4, 'D': 4.5, 'C': 5, 'B': 5.5, 'B+': 6,
            'B++': 6.5, 'A': 7, 'A+': 7.5, 'A++': 8, 'O': 8.5,
            'O+': 9, 'O++': 9.5, 'O+++': 10
        }
        
        self.grade_ranges = {
            (0, 34): 'F', (35, 39): 'E', (40, 44): 'D', (45, 49): 'C',
            (50, 54): 'B', (55, 59): 'B+', (60, 64): 'B++', (65, 69): 'A',
            (70, 74): 'A+', (75, 79): 'A++', (80, 84): 'O', (85, 89): 'O+',
            (90, 94): 'O++', (95, 100): 'O+++'
        }

    def get_grade_and_points(self, marks: float) -> Tuple[str, float]:
        """Get grade and grade points for given marks."""
        for (lower, upper), grade in self.grade_ranges.items():
            if lower <= marks <= upper:
                return grade, self.grade_points[grade]
        return 'F', 0

    def distribute_attendance_bonus(self, student_data: Dict) -> Dict:
        """Distribute attendance bonus marks optimally."""
        # Initialize variables for optimization
        prob = LpProblem("Marks_Distribution", LpMaximize)
        
        # Create variables for bonus distribution
        bonus_vars = {}
        for subject, data in student_data['subjects'].items():
            if 'theory' in data:
                bonus_vars[f'{subject}_theory'] = LpVariable(f'bonus_{subject}_theory', 0, 7)
            if 'practical' in data:
                bonus_vars[f'{subject}_practical'] = LpVariable(f'bonus_{subject}_practical', 0, 7)

        # Add constraints for total bonus marks
        prob += lpSum(bonus_vars.values()) <= student_data['attendance_bonus']

        # Add passing constraints
        for subject, data in student_data['subjects'].items():
            if 'theory' in data:
                prob += data['theory'] + bonus_vars[f'{subject}_theory'] >= 35
            if 'practical' in data:
                prob += data['practical'] + bonus_vars[f'{subject}_practical'] >= 35

        # Solve the problem
        prob.solve()

        # Extract results
        result = {}
        for subject, data in student_data['subjects'].items():
            result[subject] = {}
            if 'theory' in data:
                result[subject]['theory'] = data['theory'] + value(bonus_vars[f'{subject}_theory'])
            if 'practical' in data:
                result[subject]['practical'] = data['practical'] + value(bonus_vars[f'{subject}_practical'])

        return result

    def apply_hod_bonus(self, marks: Dict) -> Dict:
        """Apply HOD bonus marks if needed."""
        failing_subjects = []
        for subject, components in marks.items():
            if any(val < 35 for val in components.values()):
                failing_subjects.append(subject)

        if len(failing_subjects) == 1:
            subject = failing_subjects[0]
            for component in marks[subject]:
                if marks[subject][component] >= 33 and marks[subject][component] < 35:
                    marks[subject][component] = 35

        return marks

    def apply_extra_bonus(self, marks: Dict) -> Dict:
        """Apply extra bonus marks if all subjects are passed."""
        if all(all(val >= 35 for val in components.values()) 
               for components in marks.values()):
            for subject in marks:
                for component in marks[subject]:
                    marks[subject][component] += 2

        return marks

    def calculate_spi(self, final_marks: Dict, credits: Dict) -> float:
        """Calculate SPI based on final marks and credits."""
        total_grade_points = 0
        total_credits = 0

        for subject, components in final_marks.items():
            subject_marks = sum(components.values()) / len(components)
            _, grade_points = self.get_grade_and_points(subject_marks)
            
            subject_credits = sum(credits[subject].values())
            total_grade_points += grade_points * subject_credits
            total_credits += subject_credits

        return total_grade_points / total_credits if total_credits > 0 else 0