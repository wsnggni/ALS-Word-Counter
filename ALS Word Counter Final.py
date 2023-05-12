import cProfile
import docx

def myFunc():

    excluded_headers = ['References']
    excluded_figures = ['F', 'i', 'g', 'u', 'r', 'e']
    excluded_captions = ['Catapult', 'Cocked', 'Figure']

    file_type = input("Enter 'T' for .txt file or 'D' for .docx file: ")

    if file_type == 'T':
        with open('ALS_Report.txt', 'r', encoding = 'utf8') as file:
            contents = file.read()
            words = contents.split()
            skip_section = False
            print(words)
            num_words = 0
            num_figures = 0
            num_captions = 0
            for word in words:
                if word in excluded_headers:
                    skip_section = True
                    break
                if word in excluded_figures:
                    num_words -= 1
                    num_figures += 1
                if word in excluded_captions:
                    num_words -= 8
                    num_captions += 8
                if not skip_section:
                    num_words += 1
                    print(word)

        omitted_words = len(words) - num_words
        omitted_figures = num_figures//len(excluded_figures)
        print('The number of words is: ', num_words)
        print('The number of omitted words is: ', omitted_words)
        print('The number of omitted figures is: ', omitted_figures)
        print('The number of ommitted captions is: ', num_figures)

    elif file_type == 'D':
        doc_file = 'ALS_Report DOCX.docx'

        doc = docx.Document(doc_file)
        full_text = []

        for para in doc.paragraphs:
            full_text.append(para.text)

        full_text = '\n'.join(full_text)
        words = full_text.split()
        skip_section = False
        print(words)
        num_words = 0
        num_figures = 0
        num_captions = 0

        for word in words:
            if word in excluded_headers:
                skip_section = True
                break
            if word in excluded_figures:
                num_words -= 1
                num_figures += 1
            if word in excluded_captions:
                num_words -= 8
                num_captions += 8
            if not skip_section:
                num_words += 1
                print(word)

        omitted_words = len(words) - num_words
        omitted_figures = num_figures // len(excluded_figures)
        print('The number of words is: ', num_words)
        print('The number of omitted words is: ', omitted_words)
        print('The number of omitted figures is: ', omitted_figures)
        print('The number of ommitted captions is: ', num_figures)
    else:
        print('Invalid file type')


cProfile.run('myFunc()')