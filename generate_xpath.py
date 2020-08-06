import json
import pandas as pd
import random


def is_single_word(inp):
    return True if len(inp.split()) == 1 else False

def process_single_word(inp):
    return f"[text()='{inp}']"

def process_multi_word(inp):
    words = inp.split()
    xpath = f"[starts-with(text(), '{words[0]}') "
    for i in range(1, len(words)):
        if words[i] not in ["of", "and", "in", "at", "to"]:
            xpath += f"and contains(text(), '{words[i]}') "
    return xpath.strip() + "]"

def generate_xpaths(text_inp, xpath_pre, xpath_post):
    inp_list = text_inp.split(",")
    inp_list = [x.strip() for x in inp_list]
    return_list = []

    for item in inp_list:
        if is_single_word(item):
            xpath_center = process_single_word(item)
        else:
            xpath_center = process_multi_word(item)
        
        xpath = xpath_pre + xpath_center + xpath_post
        return_list.append((item, xpath))
    
    return return_list

def write_to_excel(arr, file_path, sheet_name):
    final_arr = []
    for item in arr:
        final_arr.append((item[0], "BEGIN", None, None))
        final_arr.append((None, None, "xpath", item[1]))
        final_arr.append((None, "END", None, None))
    
    df = pd.DataFrame(final_arr)

    with pd.ExcelWriter(file_path, mode='a') as writer: 
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)


if __name__ == "__main__":

    textbox = {
        'pre': '//span',
        'post': '/parent::div/following-sibling::div/input'
    }

    textarea = {
        'pre': '//span',
        'post': '/parent::div/following-sibling::div/textarea'
    }    

    combobox = {
        'pre': '//span',
        'post': '/parent::div/following-sibling::div/span/input'
    }

    datebox = {
        'pre': '//span',
        'post': '/parent::div/following-sibling::div/span/input'
    }

    checkbox = {
        'pre': '//label',
        'post': '/preceding-sibling::input'
    }

    radio = {
        'pre': '//label',
        'post': '/preceding-sibling::input'
    }       

    label = {
        'pre': '//span',
        'post': '/parent::div/following-sibling::div/span'
    }    

    button = {
        'pre': '//button',
        'post': ''
    }    
    
    with open('config.json', 'r') as f:
        config = json.load(f)   

    for item in config:
        process = item.get("process", "")
        sheet_name = item.get("sheet_name", "")

        process = "n" if process == "" else process

        sheet_name = str(random.randint(100, 1000)) if sheet_name == "" else sheet_name
        
        textbox_fields = item.get('textbox_fields', '')
        textarea_fields = item.get('textarea_fields', '')
        combobox_fields = item.get('combobox_fields', '')
        datebox_fields = item.get('datebox_fields', '')
        checkbox_fields = item.get('checkbox_fields', '')
        radio_fields = item.get('radio_fields', '')
        label_fields = item.get('label_fields', '')
        button_fields = item.get('button_fields', '')
        
        arr = []
        if (process.lower()[0]) == "y":
            if textbox_fields.strip() != "":
                arr += generate_xpaths(textbox_fields, textbox['pre'], textbox['post'])
            if textarea_fields.strip() != "":
                arr += generate_xpaths(textarea_fields, textarea['pre'], textarea['post'])                
            if combobox_fields.strip() != "":
                arr += generate_xpaths(combobox_fields, combobox['pre'], combobox['post'])
            if datebox_fields.strip() != "":
                arr += generate_xpaths(datebox_fields, datebox['pre'], datebox['post'])
            if checkbox_fields.strip() != "":
                arr += generate_xpaths(checkbox_fields, checkbox['pre'], checkbox['post'])
            if radio_fields.strip() != "":
                arr += generate_xpaths(radio_fields, radio['pre'], radio['post'])                  
            if label_fields.strip() != "":
                arr += generate_xpaths(label_fields, label['pre'], label['post'])                
            if button_fields.strip() != "":
                arr += generate_xpaths(button_fields, button['pre'], button['post'])   

        if len(arr) > 0:
            write_to_excel(arr, 'xpath.xlsx', sheet_name)                                                