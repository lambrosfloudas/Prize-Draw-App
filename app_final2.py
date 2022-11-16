import tkinter as tk
import random as rd
import pandas as pd
import datetime
from tkinter import messagebox
import tkinter.filedialog as tkfiledialog


window = tk.Tk()

window.title("Prize Winners Selection")

window.geometry('1000x1000')

participants_list = []
results_list = []
gifts = pd.read_excel('gifts_descriptions.xlsx',header=None)[0]
number_of_gift = tk.IntVar()
number_of_gift.set(0)

def participantslist():
    global number_of_gift
    global results_list

    try:
        participants_number = int(entry_field.get())
    except:
        participants_number = 0
    if participants_number >= len(gifts):
        B_next['state'] = 'normal'
        participants_input['state'] = 'disabled'
        participants_list.append(list(range(1, participants_number + 1))) 
        input_display = tk.Text(master=window, height=10, width =53, font=('Helvetica', 24))
        input_display.place(x=20, y=300)
        input_display.insert(tk.END, 'Now press the button Next to Continue')
        
    else:
        input_display = tk.Text(master=window, height=10, width =53, font=('Helvetica', 24))
        input_display.place(x=20, y=300)
        input_display.insert(tk.END, 'The participants must be at least ' + str(len(gifts)))
        #participants_input.grid(column=3, row=0)
        

    return participants_list



# Use the next button to make Draw possible again.
def next_button(number_of_gift):
    
    try:
        #button1.grid_forget() #Button to hide
        draw['state'] = 'normal'#Button to reappear
        B_next['state'] = 'disabled'

        gift_display = tk.Text(master=window, height =10, width =53, font=('Helvetica', 24))
        gift_display.place(x=20, y=300)

        gift = gifts.iloc[number_of_gift.get()]
        gift_display.insert(tk.END, 'The next present is '+gift)
    except:
        df = pd.concat(results_list).reset_index(drop=True)
        df.to_excel('results.xlsx', sheet_name='sheet1', index=False)
        finished_display = tk.Text(master=window, height=10, width =53, font=('Helvetica', 24))
        finished_display.place(x=20, y=300)
        finished_display.insert(tk.END, 'The gifts are finished')
        
    
# Draw button function.
def draw_winner_display(number_of_gift):
    # Buttons hide and reappear
    B_next['state'] = 'normal'
    draw['state'] = 'disabled'
    save['state'] = 'normal'
    #Get timestamp
    ct = datetime.datetime.now()
    try:
        random_choice_from_list = rd.choice(participants_list[0])
        gift = gifts.iloc[number_of_gift.get()]
        winner_gift = [random_choice_from_list, gift, ct]

        results_list.append(pd.DataFrame({'Winner' : winner_gift[0], 'Gift': winner_gift[1], 
                                          'Timestamp': winner_gift[2]},index=[0]))

        winner_display = tk.Text(master=window, height=10, width =53, font=('Helvetica', 24))
        winner_display.place(x=20, y=300)

        string = "The lucky number is "+str(winner_gift[0])+"\nThe participant " + str(winner_gift[0]) + ' wins ' + 'the present ' + str(winner_gift[1])
        winner_display.insert(tk.END, string)

        participants_list[0].remove(random_choice_from_list)
        number_of_gift.set(number_of_gift.get()+ 1)
    except:
        df = pd.concat(results_list).reset_index(drop=True)
        df.to_excel('results.xlsx', sheet_name='sheet1', index=False)
        finished_display = tk.Text(master=window, height=10, width =53, font=('Helvetica', 24))
        finished_display.place(x=20, y=300)
        B_next['state'] = 'disabled'
        draw['state'] = 'disabled'
        tk.messagebox.showinfo("End", "All gifts have been given!")
        finished_display.insert(tk.END, 'The gifts are finished')
        
def reset_button():
    participants_list.clear()
    results_list.clear()
    df = pd.DataFrame().reset_index(drop=True)
    number_of_gift.set(0)
    participants_input['state'] = 'normal'
    B_next['state'] = 'disabled'
    draw['state'] = 'disabled'
    save['state'] = 'disabled'
    reset_display = tk.Text(master=window, height=10, width =53, font=('Helvetica', 24))
    reset_display.place(x=20, y=300)
    reset_display.insert(tk.END, 'Please reenter the number of participants.')

def save_button():
    excel_filepath = tkfiledialog.asksaveasfilename(title='enter file name', defaultextension=".xlsx",
                                                    filetypes=(('excel files', '*.xlsx'),
                                                               ('all files', '*.*')))
    df = pd.concat(results_list)
    df.to_excel(excel_filepath, sheet_name='sheet1', index=False)
    save_display = tk.Text(master=window, height=10, width =53, font=('Helvetica', 24))
    save_display.place(x=20, y=300)
    save_display.insert(tk.END, 'The results have been saved.')

# Label
title = tk.Label(text='What is the number of participants? ', font =('Helvetica', 24))
title.place(x=30, y=30)

number_of_gift = tk.IntVar()
number_of_gift.set(0)
#Entry Field
entry_field = tk.Entry(font = ('Helvetica', 24))
entry_field.place(x=600, y=30)

#participants_input = tk.Button(text='Start', command=participantslist)
participants_input = tk.Button(text='Start', font=('Helvetica', 24), command=lambda: participantslist(), background='cyan')
participants_input.place(x=700, y= 100)

# Next Button
B_next = tk.Button(text='Next', command= lambda: next_button(number_of_gift), font=('Helvetica', 24, 'bold', 'italic'), background = 'green', state='disabled')
B_next.place(x=50, y=150)

#Draw Button
draw = tk.Button(text='Draw', command=lambda: draw_winner_display(number_of_gift), font=('Helvetica', 24, 'bold', 'italic'), background = 'green', state= 'disabled')
draw.place(x=200, y=150)
#draw.grid(column=3, row=1)

# Make a reset button.
reset = tk.Button(text='Reset', command=lambda: reset_button(), font =
('calibri', 20, 'bold', 'underline'),
                  background = 'red')
reset.place(x=800, y=800)

#Make a save button.
save = tk.Button(text='Save', command=lambda: save_button(), font =('calibri', 20, 'bold', 'underline'), background = 'orange', state= 'disabled')
save.place(x=600, y=800)
#B_next.grid(column=1, row=1)

window.mainloop()