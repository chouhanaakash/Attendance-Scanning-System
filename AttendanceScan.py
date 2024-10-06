import tkinter as tk
from tkinter import filedialog, messagebox
from paddleocr import PaddleOCR
import pandas as pd
import cv2

# Initialize PaddleOCR model
ocr = PaddleOCR(use_angle_cls=True, lang='en')

#File Path Input
def select_file(): 
    global img_path
    img_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
    if img_path:
        path_label.config(text=f"Selected Image: {img_path}")

def Image_convertor():
    if not img_path:
        messagebox.showwarning("Warning", "Please select an image file first.")
        return

    result = ocr.ocr(img_path, cls=True)
    enroll = []
    name = []
    f = False
    a = []
    temp = []
    for i in range(len(result)):
        for j in range(len(result[i])):
            data = result[i][j][1][0]
            if '0801' in data:
                if len(data) > 20:
                    name.append(data)
                enroll.append(data)
            elif len(data) > 3:
                if data == 'Name':
                    f = True
                elif f == True:
                    a.append(temp)
                    temp = []
                    name.append(data)
            elif len(data) < 2 and data in ('A', 'P'):
                temp.append(data)

    for i in range(len(enroll)):
        if enroll[i][0] != '0':
            enroll[i] = enroll[i][2:]
        if len(enroll[i]) > 12:
            enroll[i] = enroll[i][:12]
    for i in range(len(name)):
        if '0801' in name[i]:
            name[i] = name[i][12:].strip()
    attd = []
    for x in a:
        count = 0
        for y in x:
            if y == 'P':
                count += 1
        attd.append(count)

    dic = {"Enrollment": enroll, "Name": name, "Attendance": attd}
    df = pd.DataFrame(dic)

    #dialogbox for savefile name
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if excel_path:
        df.to_excel(excel_path, index=False)
        messagebox.showinfo("Success", "File created successfully!")


    dataframe_text.delete(1.0, tk.END)  
    dataframe_text.insert(tk.END, df.to_string(index=False))
    
#exit function
def exit_program():
    root.destroy()

#GUI Part
root = tk.Tk()
root.title("Attendance")
root.geometry("600x600")
root.configure(bg="#D0F0C0")  


font_style = ("Helvetica", 12)
button_style = {
    "bg": "white",   
    "fg": "blue",   
    "activebackground": "#333333",  
    "font": ("Helvetica", 12, "bold"),
    "width": 20,
    "bd": 0,
    "relief": "flat",
    
}
button_style1 = {
    "bg": "white",   
    "fg": "black",   
    "activebackground": "#333333",  
    "font": ("Helvetica", 12, "bold"),
    "width": 10,
    "bd": 0,
    "relief": "flat",
   
}
label_style = {
    "bg": "#D0F0C0",
    "fg": "#333333",
    "font": font_style
}


title_label = tk.Label(root, text="Welcome", font=("Helvetica", 16, "bold"), bg="#D0F0C0", fg="#2F4F4F")
title_label.pack(pady=20)

select_button = tk.Button(root, text="Select Image", command=select_file, **button_style)
select_button.pack(pady=10)
submit_button = tk.Button(root, text="Submit", command=Image_convertor, **button_style)
submit_button.pack(pady=10)

path_label = tk.Label(root, text="No Image Selected", **label_style)
path_label.pack(pady=5)


dataframe_text = tk.Text(root, height=20, width=60, wrap='none', font=("Arial", 12))
dataframe_text.pack(pady=20)

footer_label = tk.Label(root, text="Developed by Aakash CHouhan", font=("Helvetica", 10), bg="#D0F0C0", fg="#2F4F4F")
footer_label.pack(side="bottom", pady=10)


exit_button = tk.Button(root, text="Close", command=exit_program, **button_style1)
exit_button.pack(side="bottom", pady=10)


root.mainloop()