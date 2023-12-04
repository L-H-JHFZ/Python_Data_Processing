import tkinter as tk  
import time  
  
def time_now():  
    current_time = time.strftime('%H:%M:%S')  
    clock.config(text=current_time)  
    clock.after(1000, time_now)  
  
root = tk.Tk()  
clock = tk.Label(root, font=('times', 50, 'bold'), bg='green')  
clock.pack(fill='both', expand=1)  
time_now()  
root.mainloop()
