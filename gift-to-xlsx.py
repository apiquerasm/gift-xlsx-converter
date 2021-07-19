import argparse
from pygiftparser import parser
import xlsxwriter   # import xlsxwriter module

def parse_input_arguments():
    parser = argparse.ArgumentParser(description='GIFT Moodle to XLSX Kahoot converter')
    parser.add_argument('-f', '--file', dest='file', required=True,
                        help='GIFT Moodle file.')
    args = parser.parse_args()
    return args

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

if __name__ == "__main__":
    QUESTION_TIME = 30

    print("·--------------------------------------·")
    print("| GIFT Moodle to XLSX Kahoot converter |")
    print("·--------------------------------------·")

    args = parse_input_arguments()
    input_file = args.file
    output_file = input_file + ".xlsx"

    print("* Input file: " + input_file)
    print("* Output file: " + output_file)

    with open(input_file, 'r') as myfile:
        s = myfile.read()
        # Parse GIFT file:
        result = parser.parse(s)

        # Create xlsx document
        workbook = xlsxwriter.Workbook(output_file)
        # Create a new sheet in xlsx document
        worksheet = workbook.add_worksheet()
        
        # Cell format for header
        format_header = workbook.add_format({'bold': True, 'font_color': 'blue'})
        format_bold = workbook.add_format({'bold': True})

        # Write Kahoot header template
        worksheet.write("B2", "Kahoot! Quiz Template", format_header)
        worksheet.write("B3", "Add questions, at least two answer alternatives, time limit and choose correct answers (at least one). Have fun creating your awesome quiz!")
        worksheet.write("B4", "Remember: questions have a limit of 120 characters and answers can have 75 characters max. Text will turn red in Excel or Google Docs if you exceed this limit. If several answers are correct, separate them with a comma.")
        worksheet.write("B5", "See an example question below (don't forget to overwrite this with your first question!)")
        worksheet.write("B6", "And remember,  if you're not using Excel you need to export to .xlsx format before you upload to Kahoot!")
        worksheet.write("B8", "Question - max 120 characters", format_bold)
        worksheet.write("C8", "Answer 1 - max 75 characters", format_bold)
        worksheet.write("D8", "Answer 2 - max 75 characters", format_bold)
        worksheet.write("E8", "Answer 3 - max 75 characters", format_bold)
        worksheet.write("F8", "Answer 4 - max 75 characters", format_bold)
        worksheet.write("G8", "Time limit (sec) – 5, 10, 20, 30, 60, 90, 120, or 240 secs", format_bold)
        worksheet.write("H8", "Correct answer(s) - choose at least one", format_bold)

        # Set column width
        worksheet.set_column('A:A', 5)
        worksheet.set_column('B:B', 60)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 20)
        worksheet.set_column('F:F', 20)
        worksheet.set_column('G:G', 10)
        worksheet.set_column('H:H', 10)

        # Start from the 8th cell. Rows and columns are zero indexed.
        row = 8

        # Input file question counter
        num = 1

        # Compatible questions counters
        question_index = 0

        for question in result.questions:
            question_text = question.text
            question_type = repr(question.answer)

            print("\n" + str(num) + ". " + question_text)
            print("Type: " + question_type)

            if question_type == 'TrueFalse()':
                question_index += 1
                print(f"{bcolors.OKGREEN}OK. Imported{bcolors.ENDC}")

                # Add question to xlsx document  
                worksheet.write(row, 0, question_index)
                worksheet.write(row, 1, question_text)
                worksheet.write_string(row, 2, 'Verdadero')
                worksheet.write_string(row, 3, 'Falso')

                # Write question time
                worksheet.write(row, 6, QUESTION_TIME)
                # Write correct answer
                if question.answer.options[0].text == 'True':
                    worksheet.write(row, 7, 1)
                else:
                    worksheet.write(row, 7, 2)
                row += 1

            elif question_type == 'MultipleChoiceRadio()':
                question_index += 1
                print(f"{bcolors.OKGREEN}OK. Imported{bcolors.ENDC}")

                # Add question to xlsx document
                worksheet.write(row, 0, question_index)
                worksheet.write(row, 1, question_text)

                for i in range(0,4):
                    # Write an anwser
                    worksheet.write(row, i+2, question.answer.options[i].text)
                    # Write correct answer
                    if question.answer.options[i].prefix == '=':
                        worksheet.write(row, 7, i+1) # ERROR aparece un ' antes del número

                # Write question time
                worksheet.write(row, 6, QUESTION_TIME)
                row += 1
            else:
                print(f"{bcolors.FAIL}ERROR: No compatible{bcolors.ENDC}")

            # Increase question counter
            num += 1

        print("\n* Total questions: " + str(len(result.questions)))
        print("* Generated '" + output_file + "' with " + str(question_index) + " compatible questions.")
        print("\n")
        # Finally, close the Excel file
        # via the close() method.
        workbook.close()

