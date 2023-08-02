import uuid

from excel2pycl import Cell, Parser


def main():
    translation_file_path = f'temp/{uuid.uuid4()}.py'

    Parser() \
        .set_excel_file_path('./test.xlsx') \
        .enable_safety_check() \
        .write_translation(translation_file_path)

    print(translation_file_path)


if __name__ == '__main__':
    main()


