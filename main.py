import sys
from excel_operations import ExcelHandler

def main():
    if len(sys.argv) != 3:
        print("Usage: python main.py <input_excel_file> <output_excel_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    try:
        # Initialize Excel handler
        handler = ExcelHandler(input_file, output_file)

        # Read and process input data
        print("Reading input data...")
        student_data = handler.read_input_data()

        # Process marks and calculate SPI
        print("Processing marks and calculating SPI...")
        final_marks, spi = handler.process_student_data(student_data)

        # Write results to output file
        print("Writing results to output file...")
        handler.write_output(final_marks, spi)

        print(f"\nProcessing complete! Results written to {output_file}")
        print(f"Final SPI: {spi:.2f}")

    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()