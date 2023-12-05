import subprocess
import format_phonenumber_output_excel
import read_excel_send_email

def read_excel_and_send_emails():
    read_excel_send_email.read_excel_and_send_emails(excel_file_path, html_template_path)

def format_phone_number_script():
    format_phonenumber_output_excel.concat_phone_numbers(excel_file_path, output_excel_path)

def main():
    while True:
        print("\nOptions:")
        print("1 - Enviar email")
        print("2 - Formatar numero de celular")
        print("0 - Exit")

        choice = input("Select an option: ")

        if choice == "1":
            read_excel_and_send_emails()
        elif choice == "2":
            format_phone_number_script()
        elif choice == "0":
            print("Exiting program. Goodbye!")
            break
        else:
            print("Invalid option. Please enter a valid option.")

if __name__ == "__main__":
    excel_file_path = "files/041611.xlsx"
    output_excel_path = "files/result.xlsx"
    html_template_path = "assets/html/index.html"
    main()
