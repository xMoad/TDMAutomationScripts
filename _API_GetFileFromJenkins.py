import urllib.request

url = "https://example.com/fichier.zip"
nom_fichier = "fichier_telecharge.zip"

# Envoyer une requête GET pour télécharger le fichier
reponse = urllib.request.urlopen(url)

# Ouvrir un fichier local en mode écriture binaire
with open(nom_fichier, "wb") as fichier:
    # Écrire le contenu du fichier téléchargé dans le fichier local
    fichier.write(reponse.read())

print(f"Téléchargement de {url} terminé dans {nom_fichier}")