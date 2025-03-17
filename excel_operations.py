import pandas as pd
from typing import Dict, List, Tuple
from mark_distribution import MarkDistributor

class ExcelHandler:
    def __init__(self, input_file: str, output_file: str):
        self.input_file = input_file
        self.output_file = output_file
        self.mark_distributor = MarkDistributor()

    def read_input_data(self) -> Dict:
        """Read student data from input Excel file."""
        try:
            # Read the Excel file
            df = pd.read_excel(self.input_file)
            
            # Initialize student data structure
            student_data = {
                'subjects': {},
                'attendance_bonus': float(df.loc[df['Parameter'] == 'Attendance Bonus', 'Value'].iloc[0])
            }

            # Process subject data
            subjects_df = df[df['Parameter'].str.contains('Subject', na=False)]
            for _, row in subjects_df.iterrows():
                subject_name = row['Subject Name']
                student_data['subjects'][subject_name] = {}
                
                if not pd.isna(row['Theory Marks']):
                    student_data['subjects'][subject_name]['theory'] = float(row['Theory Marks'])
                if not pd.isna(row['Practical Marks']):
                    student_data['subjects'][subject_name]['practical'] = float(row['Practical Marks'])

            # Read credits
            student_data['credits'] = {}
            for subject_name in student_data['subjects']:
                student_data['credits'][subject_name] = {}
                subject_row = df[df['Subject Name'] == subject_name].iloc[0]
                if 'theory' in student_data['subjects'][subject_name]:
                    student_data['credits'][subject_name]['theory'] = float(subject_row['Theory Credits'])
                if 'practical' in student_data['subjects'][subject_name]:
                    student_data['credits'][subject_name]['practical'] = float(subject_row['Practical Credits'])

            return student_data

        except Exception as e:
            raise Exception(f"Error reading input file: {str(e)}")

    def process_student_data(self, student_data: Dict) -> Tuple[Dict, float]:
        """Process student data and return final marks and SPI."""
        # Distribute attendance bonus
        marks_with_bonus = self.mark_distributor.distribute_attendance_bonus(student_data)
        
        # Apply HOD bonus if needed
        marks_with_hod = self.mark_distributor.apply_hod_bonus(marks_with_bonus)
        
        # Apply extra bonus if applicable
        final_marks = self.mark_distributor.apply_extra_bonus(marks_with_hod)
        
        # Calculate final SPI
        spi = self.mark_distributor.calculate_spi(final_marks, student_data['credits'])
        
        return final_marks, spi

    def write_output(self, final_marks: Dict, spi: float) -> None:
        """Write results to output Excel file."""
        try:
            # Create output dataframe
            output_data = []
            for subject, components in final_marks.items():
                row = {'Subject': subject}
                row.update(components)
                output_data.append(row)

            results_df = pd.DataFrame(output_data)
            
            # Add SPI information
            spi_df = pd.DataFrame([{'Parameter': 'Final SPI', 'Value': spi}])
            
            # Write to Excel
            with pd.ExcelWriter(self.output_file) as writer:
                results_df.to_excel(writer, sheet_name='Final Marks', index=False)
                spi_df.to_excel(writer, sheet_name='SPI', index=False)

        except Exception as e:
            raise Exception(f"Error writing output file: {str(e)}")