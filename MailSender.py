import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from tkinter import messagebox, filedialog
import tkinter as tk
from typing import Tuple
import time
import customtkinter as ctk

class ConnexionPage(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Connexion")
        self.geometry("200x320")
        self.resizable(width=False, height=False)
        self.widgets()

    def widgets(self):
        ctk.CTkLabel(self, text="Adresse email:").pack()
        self.email_entry = ctk.CTkEntry(self)
        self.email_entry.pack()
        ctk.CTkLabel(self, text="Mot de passe:").pack()
        self.password_entry = ctk.CTkEntry(self)
        self.password_entry.pack()
        ctk.CTkLabel(self, text="Serveur:").pack()
        self.server_entry = ctk.CTkEntry(self)
        self.server_entry.pack()
        ctk.CTkLabel(self, text="Port:").pack()
        self.port_entry = ctk.CTkEntry(self)
        self.port_entry.pack()
        self.connexionButton = ctk.CTkButton(self, text="Connexion", command=self.connexion)
        self.connexionButton.pack(pady=20)

    def connexion(self):
        try:
            smtp_server = smtplib.SMTP(self.server_entry.get(), self.port_entry.get())
            smtp_server.ehlo()
            smtp_server.starttls()
            smtp_server.login(self.email_entry.get(), self.password_entry.get())
            print("Connexion réussie !")
            ctk.CTk.destroy
            time.sleep(0.5)
            application = EmailApp(self.email_entry.get(), self.password_entry.get(), self.server_entry.get(), self.port_entry.get())
            application.mainloop()
        except Exception as e:
            print(e)
            messagebox.showwarning("Erreur", f"Connexion impossible: {e}")

class EmailApp(ctk.CTk):

    def __init__(self, email, password, server, port):
        super().__init__()
        self.title("Envoi d'email")
        self.geometry("500x600")
        self.resizable(width=False, height=False)
        self.sender_email = email
        self.password = password
        self.smtp_server = server
        self.port = port
        self.receiver_emails = []
        self.create_widgets()
        
    def load_receiver_emails(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        json_path = os.path.join(current_dir, "destinataires.json")
        try:
            with open(json_path, "r") as json_file:
                data = json.load(json_file)
                return data.get("destinataires", [])  # Retourner les emails chargés depuis le fichier
        except FileNotFoundError:
            messagebox.showwarning("Avertissement", "Le fichier 'destinataires.json' n'a pas été trouvé.")
            return []

    def create_widgets(self):
        ctk.CTkLabel(self, text='Destinataire(s):').pack()
        self.receiver_emails = self.load_receiver_emails()  # Charger les destinataires
        self.receiver_combobox = ctk.CTkComboBox(self, values=self.receiver_emails, width=200)
        self.receiver_combobox.pack()

        ctk.CTkLabel(self, text='Sujet:').pack()
        self.subject_entry = ctk.CTkEntry(self, width=400, height=40)
        self.subject_entry.pack()

        ctk.CTkLabel(self, text="Message:").pack()
        self.body_entry = ctk.CTkTextbox(self, width=400, height=300)
        self.body_entry.pack(side="top")

        self.attach_button = ctk.CTkButton(self, text='Pièce jointe', command=self.attach_file)
        self.attach_button.pack(pady=10)

        self.attachment_label = ctk.CTkLabel(self, text='')
        self.attachment_label.pack()

        self.send_button = ctk.CTkButton(self, text='Envoyer', command=self.send_email)
        self.send_button.pack(pady=20)

        self.attachment_path = None

    def attach_file(self):
        self.attachment_path = filedialog.askopenfilename()
        if self.attachment_path:
            self.attachment_label.configure(text=os.path.basename(self.attachment_path))

    def send_email(self):
        receiver_emails = self.receiver_combobox.cget('values')  # Utiliser cget pour obtenir les valeurs de la combobox
        subject = self.subject_entry.get()
        body = self.body_entry.get("1.0", "end-1c")
        if not receiver_emails or not subject or not body:
            messagebox.showerror('Erreur', 'Veuillez remplir tous les champs.')
            return

        for receiver_email in receiver_emails:
            message = MIMEMultipart()
            message['From'] = self.sender_email
            message['To'] = receiver_email
            message['Subject'] = subject
            message.attach(MIMEText(body, 'plain'))

            if self.attachment_path:
                attachment = open(self.attachment_path, "rb")
                part = MIMEBase('application', 'octet-stream')
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', "attachment; filename= " + os.path.basename(self.attachment_path))
                message.attach(part)

            try:
                with smtplib.SMTP(self.smtp_server, self.port) as server:
                    server.starttls()
                    server.login(self.sender_email, self.password)
                    text = message.as_string()
                    server.sendmail(self.sender_email, receiver_email, text)
                messagebox.showinfo('Succès', 'Email envoyé avec succès !')
            except Exception as e:
                messagebox.showerror('Erreur', f"Une erreur s'est produite : {str(e)}")



if __name__ == '__main__':
    current_dir = os.path.dirname(__file__)
    json_file = os.path.join(current_dir, "destinataires.json")
    if not os.path.exists(json_file):
        data = {
            "destinataires": [""]
        }
        with open(json_file, "w") as f:
            json.dump(data, f)
            print("Fichier créé avec succès")

    app = ConnexionPage()
    app.mainloop()
