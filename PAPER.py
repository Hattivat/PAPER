# -*- coding: utf-8 -*-
"""
Created on Thu Mar 10 17:29:51 2016

@author: Jan Janiczek
"""

import tkinter
from contextlib import redirect_stdout
import inboxcleaner
import moveIT
import passgen


class Std_redirector(object):
    """This elegant solution has been shamelessly copied from the internet."""

    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        self.widget.insert(tkinter.END, string)
        self.widget.see(tkinter.END)

if __name__ == "__main__":
    root = tkinter.Tk()
    root.title("PAPER - the PAP Email Responder")

    # Aesthetics inspired by the timeless ugliness of OS/400 UI.
    button1 = tkinter.Button(root, text="Move NoReply mail to the Untitled folder",
                            command=lambda: inboxcleaner.clean_mail(), padx=50,
                            pady=15, bg="black", fg="lime")
    button1.grid(sticky="nswe")

    button2 = tkinter.Button(root, text="Send MoveIT emails", command=lambda:
                        moveIT.main(), padx=50, pady=15, bg="black", fg="lime")
    button2.grid(sticky="nswe")

    button3 = tkinter.Button(root, text="Generate a (pseudo)random password",
                             command=lambda: passgen.pass_gen(), padx=50,
                             pady=15, bg="black", fg="lime")
    button3.grid(sticky="nswe")

    textbox = tkinter.Text(root, height=5, width=30, bg="black", fg="lime")
    textbox.grid(sticky="we")

    # Redirects the output of all "print" calls in other functions to the textbox.
    yolo = Std_redirector(textbox)

    with redirect_stdout(yolo):
        root.mainloop()
