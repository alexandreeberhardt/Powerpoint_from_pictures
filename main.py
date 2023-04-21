from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import os

def powerpoint_from_dossier(chemin_du_dossier):
    #liste toutes les images dans le dossier cité
    liste_fichiers = os.listdir(chemin_du_dossier)
    #les ajoutes au powerpoint en conservant les proportions et sans dépasser la taille de la slide
    for img_path in liste_fichiers:
        img_path=chemin_du_dossier+"\\"+img_path
        
        # Créez une nouvelle diapositive
        slide = prs.slides.add_slide(prs.slide_layouts[0])

        # Charger l'image et récupérer ses dimensions
        img = Image.open(img_path)
        img_width_px, img_height_px = img.size

        # Calculer le ratio de la taille de l'image par rapport à la taille de la diapositive
        img_ratio = img_width_px / img_height_px
        slide_ratio = prs.slide_width / prs.slide_height

        # Redimensionner l'image en fonction de la taille de la diapositive
        if img_ratio > slide_ratio:
            # l'image est plus large que la diapositive, ajuster la largeur
            img_width = prs.slide_width
            img_height = img_width / img_ratio
        else:
            # l'image est plus haute que la diapositive, ajuster la hauteur
            img_height = prs.slide_height
            img_width = img_height * img_ratio

        # Ajouter l'image à la diapositive
        left = (prs.slide_width - img_width) / 2
        top = (prs.slide_height - img_height) / 2
        pic = slide.shapes.add_picture(img_path, left, top, width=img_width, height=img_height)
        # Insérez l'image sur la diapositive

prs = Presentation()
#lister les dossiers d'image puis iterer dessus
chemins=[".\dossier1",".\dossier2",".\dossier3"]
for chemin in chemins :
    powerpoint_from_dossier(chemin)
#enregistrement du powerpoint format pptx
prs.save('presentation.pptx')
