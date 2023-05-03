# -*- coding: utf-8 -*-
"""
Created on Sat Oct  8 16:48:35 2022

@author: le-to
"""

import pandas as pd
import streamlit as st
from pathlib import Path
from docxtpl import DocxTemplate
import convertapi
convertapi.api_secret = 'M9EovZhpcl31DxWP'
import os




st.set_page_config(page_title="Gestionnaire des attestations", page_icon=":bar_chart:", layout='wide')
footer="""<style>
a:link , a:visited{
color: grey;
background-color: transparent;

}

a:hover,  a:active {
color: red;
background-color: transparent;

}

.footer {
position: fixed;
left: 0;
bottom: 0;
width: 100%;
background-color: white;
color: tomato;
text-align: center;
}
</style>
<div class="footer">
<p>E-Certificate Auto Generator For Scientific Events  <a style='display: block; text-align: center;' href="mailto:hadmaher83@gmail.com" target="_blank"> 🏆  Geni-us © 2022</a></p>
</div>
"""
st.markdown(footer,unsafe_allow_html=True)



filelist=[]
path = Path(__file__).parent
for (root, dirs, file) in os.walk(path):
    for f in file:
        if '.xlsx' in f or '.XLSX' in f:
                  filename=os.path.basename(f).split(".")[0]
                  filelist.append(filename)
        
congres_selected=st.sidebar.selectbox("Sélectionner le congrès souhaité:", options=filelist)               


form=st.sidebar.form('generate')
mail=form.text_input("Entrer le mail du Premier Auteur", key='mail')
envoyer = form.form_submit_button('Envoyer')

if "envoyer_state" not in st.session_state:
    st.session_state.envoyer_state = False

if envoyer or st.session_state.envoyer_state : 
    st.session_state.envoyer_state= True
        
    excel_file = congres_selected+'.xlsx'
    
    
    df = pd.read_excel(excel_file, sheet_name='DATA', usecols='A:F', header=0 ) 
    df_o = pd.read_excel(excel_file, sheet_name='ORALE', usecols='A:F', header=0 )
    df_mail=df.query("Email == @mail")
    df_o_mail=df_o.query("Email == @mail")
    nbr=len(df_mail)
    nbr_o=len(df_o_mail)
    st.info("Vous avez à votre compte:") 
    st.success(f"{nbr} attestation(s) pour communication affichée(s)")
    st.success(f"{nbr_o} attestation(s) pour communication Orale(s)")
    #df['Auteur']=df['Nom']+ " "+ df['Prenom']
    tab1, tab2 = st.tabs(["Communication Affichée", "Communication Orale"])
    with tab1:
              
        
        if nbr > 0:
                
            
            option_titre=df_mail.Titre.unique().tolist()
            titre_select=st.selectbox("Selectionner la communication Affichée",  option_titre, key='sel' )
            st.write(f"Vous avez selectionné: {titre_select}")
            df_com=df_mail.query("Titre == @titre_select")
            autres_auteurs=df_com['Collaborateurs'].unique().tolist()
            Auteur=df_com['Premier Auteur'].unique().tolist()
            
            for val in autres_auteurs:
                autres_auteurs=val
            for val in Auteur:
                Auteur=val
            
            base_dir = Path(__file__).parent
            script = congres_selected+".DOCX"
            word_template_path = base_dir / script
            output_dir = base_dir / "certificate/docx/affiche"
            output_pdf = base_dir / "certificate/certificatepdf/affiche"
            doc = DocxTemplate(word_template_path)
                    
        
            context = {
                "Auteur" : Auteur,
                "Titre" : titre_select,
                "Collab" : autres_auteurs
                }
        
            doc.render(context)
            output_path = output_dir / "certificate.docx"
            doc.save(output_path)
            saved_file = output_dir
            converted = convertapi.convert('pdf', { 'File': './certificate/docx/affiche/certificate.docx' })
            converted.file.save('./certificate/certificatepdf/affiche/certificate.pdf')

            
            
            certificate = "certificate.pdf"
            filepath = output_pdf/ certificate
            
            with open(filepath, "rb") as pdf_file:
                PDFbyte = pdf_file.read()
            st.download_button(label="Telecharger votre Attestation", 
                                   data=PDFbyte,
                                   file_name="certificate.pdf",
                                   mime='application/octet-stream')
        else:
            st.error("vous n'avez pas participé à ce congrès - Vérifier le mail du premier auteur")
            
        with tab2:
            
            if nbr > 0:
                    
                
                option_titre1=df_o_mail.Titre.unique().tolist()
                titre_select1=st.selectbox("Selectionner la communication Orale",  option_titre1, key='sel2' )
                st.write(f"Vous avez selectionné: {titre_select1}")
                df_com=df_o_mail.query("Titre == @titre_select1")
                autres_auteurs1=df_com['Collaborateurs'].unique().tolist()
                Auteur1=df_com['Premier Auteur'].unique().tolist()
                
                for val in autres_auteurs1:
                    autres_auteurs1=val
                for val in Auteur1:
                    Auteur1=val
                
                base_dir = Path(__file__).parent
                script = congres_selected+"-Orale"+".DOCX"
                word_template_path = base_dir / script
                output_dir = base_dir / "certificate/docx/orale"
                output_pdf = base_dir / "certificate/certificatepdf/orale"
                doc = DocxTemplate(word_template_path)
                        
            
                context = {
                    "Auteur" : Auteur1,
                    "Titre" : titre_select1,
                    "Collab" : autres_auteurs1
                    }
            
                doc.render(context)
                output_path = output_dir / "certificate.docx"
                doc.save(output_path)
                saved_file = output_dir
                converted = convertapi.convert('pdf', { 'File': './certificate/docx/orale/certificate.docx' })
                converted.file.save('./certificate/certificatepdf/orale/certificate.pdf')
                
                certificate = "certificate.pdf"
                filepath = output_pdf/ certificate
                
                with open(filepath, "rb") as pdf_file:
                    PDFbyte = pdf_file.read()
                st.download_button(label="Telecharger votre Attestation", 
                                       data=PDFbyte,
                                       file_name="certificate.pdf",
                                       mime='application/octet-stream')
            else:
                st.error("vous n'avez pas participé à ce congrès - Vérifier le mail du premier auteur")



    
  