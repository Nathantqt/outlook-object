import win32com.client

def chercher_messages_outlook(objet_recherche):
    # Connexion à l'application Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Accès à la boîte de réception (dossier numéro 6 correspond à "Inbox")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items

    messages_trouves = []

    # Cherche les messages dont l'objet commence par ou contient l'entrée de l'utilisateur
    for message in messages:
        if objet_recherche.lower() in message.Subject.lower():  # Contient l'objet
            messages_trouves.append(message)

    # Affiche les messages trouvés ou un message approprié si aucun n'est trouvé
    if messages_trouves:
        print(f"\nMessages trouvés contenant '{objet_recherche}':")
        for msg in messages_trouves:
            print("Objet : ", msg.Subject)
            print("Contenu :\n", msg.Body)
            print("-" * 50)  # Ligne de séparation entre les messages
    else:
        print("Aucun message trouvé contenant :", objet_recherche)

# Demander à l'utilisateur de fournir le mot clé
objet = input("Entrez le mot clé de l'objet du message à rechercher : ")

# Appeler la fonction pour chercher les messages
chercher_messages_outlook(objet)
