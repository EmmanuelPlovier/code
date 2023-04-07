import pandas as pd
import numpy as np

#C:\Perso\Carrefour\code mouvement user exit version_V10.xlsx

file_name = input('Entrez le nom du fichier contenant la table des catégories:')

df = pd.read_excel(file_name, sheet_name='Table Category')		# chargement des infos category
matrice = pd.read_excel(file_name, sheet_name='Matrice')		# chargement de la matrice

print("Nombre de matrice à traiter : " + str(len(matrice)))
print("Nombre de catégorie à traiter : " + str(len(df)))

lst_cat = df["Category"].unique()								# création d'un array numpy des catégories en fonction de la table des catégories chargées
lst_category = pd.DataFrame(lst_cat, 
                        columns = ["category"])					# chargement des catégories de numpy dans un dataframe


v_estimate = len(matrice[matrice.Action == "All"])*len(lst_category)
v_estimate2 = len(matrice)-len(matrice[matrice.Action == "All"])
v_estimation = (v_estimate+v_estimate2)*17

print("Nombre maximum de combinaison : " + str(v_estimation), "Nombre estimé : "+str(round(v_estimation*0.6)))


user_exit = pd.DataFrame ({
	"cle": ["string"], "cle_exception": ["string"], "Plan_evaluation": ["string"], "Code_mouvement": ["string"],"Categorie": ["string"],"Compte": [1],"S/L_Partenaire": ["string"],
	"Code_flux": ["string"],"Nature_retraitement": ["string"],
	})

user_exit['Compte'] = user_exit['Compte'].apply(np.int64)		# changement du type de compte en int64 pour éviter les .0 (ex : 5.0)

#lst = ["Compte Valeur Brut Aux. PCO (Drive)","Compte Amort. Lin.PCO (Drive)","Compte Dotation Lin. PCO (Drive)","Compte Amts dérogatoire PCO (Drive)",
#"Compte PCO Dotation Déro. (Drive)","Compte PCO Reprise Déro. (Drive)", "Compte PCO Reliquat Déro. (Drive)", "Compe PCO VNC Cessions PS",
# "Compe PCO Mise au rebut R et P", "Compe PCO Produit de cession", "Compe PCO Valeur Brute Aux. Impairment", "Compte PCO dotation Impairment ",
#  "Compte provision dépréciation actif PCO",  "Compte PCO dotation provision depreciation actif courant", "Compte PCO reprise provision depreciation actif courant",
#  "Compte PCO dotation provision depreciation actif exceptionnel ",  "Compte PCO reprise provision depreciation actif exceptionnel" ]

df.fillna("",inplace=True)										# mise à "" des valeurs en NAN de la table des catégories qui contient les natures
matrice.fillna("",inplace=True)									# mise à "" des valeurs en NAN de la table de la matrice

def add_code (a, b, c, d, e, f, g) : 							# création de la fonction add_code
	v_cle = str(b) + "|" + c + "|" + str(d) + "|" + g + "|" + e
	v_cle_exception = str(b) + "|" + c + "|" + str(d) + "|"  + e
	if str(d)[0] =="6" or str(d)[0] == "7"  :
		do_nothing = ""
	else :  
		user_exit.loc[len(user_exit.index)] = [v_cle,	#clé
												v_cle_exception, #clé pour la recherche des Actions exceptions
												a, 		#"Plan"
												b,  	#code mouvement 
												c, 		#categorie
												d, 		#compte
												e, 		#S/L
												f, 		#code_flux
												g 		#nature
												]
	return 1


def find_flux (compte, flux_bil, flux_pnl) :					# function pour retourner le type de flux en fonction du compte et 2 flux passé bilan et P&L
	v_temp_cpt = str(compte)[0]									# retourne le flux 
	if v_temp_cpt == "1" or v_temp_cpt == "2" or v_temp_cpt == "3" or v_temp_cpt == "4":
		v_flux = flux_bil
	else:
		v_flux = flux_pnl

	return v_flux

for x in matrice.index :										#looping sur la matrice

	v_code_mouvement = matrice.at[x, "Code_mouvement"]
	v_libelle = matrice.at[x, "Libellé"]
	v_action = matrice.at[x, "Action"]
	v_sous_category = matrice.at[x, "Sous_categorie"]
	v_compte = matrice.at[x, "Compte"]
	v_sl = matrice.at[x, "SL"]
	v_code_flux_bilan = matrice.at[x, "Code_flux_Bilan"]
	v_code_flux_pnl = matrice.at[x, "Code_Flux_P&L"]
	v_zone_produit = matrice.at[x, "Zone Product Category"]
	v_zone_amort_lin = matrice.at[x, "Zone Product amortissement linéaire"]
	v_zone_dot_lin = matrice.at[x, "Zone Product dotation linéaire PS"]
	v_zone_amort_dero = matrice.at[x, "Zone Product amortissement dérogatoire PS"]
	v_zone_dot_dero = matrice.at[x, "Zone Product dotation dérogatoire PS"]
	v_zone_reprise_dero= matrice.at[x, "Zone Product reprise dérogatoire PS"]
	v_zone_reliquat_dero = matrice.at[x, "Zone Product reliquat dérogatoire PS"]
	v_zone_vnc_cession = matrice.at[x, "Zone Product VNC Cession PS"]
	v_zone_vnc_mar = matrice.at[x, "Zone Product VNC Mise au rebut R et P"]
	v_zone_produit_cession = matrice.at[x, "Zone Product Produit cession PS"]
	v_zone_valeur_impairment = matrice.at[x, "Zone Product Valeur Brut Impairment"]
	v_zone_dot_impairment = matrice.at[x, "Zone Product Dotation Impairment"]
	v_zone_prov_dep = matrice.at[x, "Zone Product Provision dépréciation Actif"]
	v_zone_dot_prov_actif_courant = matrice.at[x, "zone product dot prov actif courrant"]
	v_zone_reprise_prov_actif_courant = matrice.at[x, "zone product rep, prov dep,actif courrant"]
	v_zone_prov_dep_excep = matrice.at[x, "zone product prov depr actif exceptionnelle"]
	v_zone_reprise_prov_dep_excep = matrice.at[x, "zone product rer, prov, dep, actif exceptionnelle"]


	if v_action == "All" :										#si on a mis All comme action

		for ind in df.index:									# looping sur le fichier des category 

			v_category = df.at[ind, "Category"]

			if df.at[ind, "Compte Valeur Brut Aux. PCO (Drive)"] != "" and v_zone_produit == "X":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte Valeur Brut Aux. PCO (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte Valeur Brut Aux. PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product Category"])

			if df.at[ind, "Compte Valeur Brut Aux. PCO (Drive)"] != "" and v_zone_produit != "X" and v_zone_produit !="":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte Valeur Brut Aux. PCO (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte Valeur Brut Aux. PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_produit)

			if df.at[ind, "Compte Amort. Lin.PCO (Drive)"] != "" and v_zone_amort_lin == "X" and v_zone_produit !="":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte Amort. Lin.PCO (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte Amort. Lin.PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product amortissement linéaire"])

			if df.at[ind, "Compte Amort. Lin.PCO (Drive)"] != "" and v_zone_amort_lin != "X" and v_zone_amort_lin != "":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte Amort. Lin.PCO (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte Amort. Lin.PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_amort_lin)

			if df.at[ind, "Compte Dotation Lin. PCO (Drive)"] != "" and v_zone_dot_lin =="X":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte Dotation Lin. PCO (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte Dotation Lin. PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product dotation linéaire PS"])

			if df.at[ind, "Compte Dotation Lin. PCO (Drive)"] != "" and v_zone_dot_lin !="X" and v_zone_dot_lin !="":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte Dotation Lin. PCO (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte Dotation Lin. PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_dot_lin)

			if df.at[ind, "Compte Amts dérogatoire PCO (Drive)"] != ""  and v_zone_amort_dero =="X":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte Amts dérogatoire PCO (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte Amts dérogatoire PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product amortissement dérogatoire PS"])

			if df.at[ind, "Compte Amts dérogatoire PCO (Drive)"] != ""  and v_zone_amort_dero !="X" and v_zone_amort_dero !="":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte Amts dérogatoire PCO (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte Amts dérogatoire PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_amort_dero)

			if df.at[ind, "Compte PCO Dotation Déro. (Drive)"] != "" and v_zone_dot_dero =="X":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO Dotation Déro. (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte PCO Dotation Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product dotation dérogatoire PS"])

			if df.at[ind, "Compte PCO Dotation Déro. (Drive)"] != "" and v_zone_dot_dero !="X" and v_zone_dot_dero !="":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO Dotation Déro. (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte PCO Dotation Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_dot_dero)

			if df.at[ind, "Compte PCO Reprise Déro. (Drive)"] != "" and v_zone_reprise_dero == "X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO Reprise Déro. (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte PCO Reprise Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product reprise dérogatoire PS"])

			if df.at[ind, "Compte PCO Reprise Déro. (Drive)"] != "" and v_zone_reprise_dero != "X" and v_zone_reprise_dero != "" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO Reprise Déro. (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte PCO Reprise Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_reprise_dero)

			if df.at[ind, "Compte PCO Reliquat Déro. (Drive)"] != "" and v_zone_reliquat_dero =="X":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO Reliquat Déro. (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte PCO Reliquat Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product reliquat dérogatoire PS"])

			if df.at[ind, "Compte PCO Reliquat Déro. (Drive)"] != "" and v_zone_reliquat_dero !="X" and v_zone_reliquat_dero !="":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO Reliquat Déro. (Drive)"], v_sl, 
					find_flux(df.at[ind, "Compte PCO Reliquat Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_reliquat_dero)

			if df.at[ind, "Compe PCO VNC Cessions PS"] != "" and v_zone_vnc_cession == "X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compe PCO VNC Cessions PS"], v_sl, 
					find_flux(df.at[ind, "Compe PCO VNC Cessions PS"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product VNC Cession PS"])

			if df.at[ind, "Compe PCO VNC Cessions PS"] != "" and v_zone_vnc_cession != "X" and v_zone_vnc_cession != "" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compe PCO VNC Cessions PS"], v_sl, 
					find_flux(df.at[ind, "Compe PCO VNC Cessions PS"],v_code_flux_bilan,v_code_flux_pnl), v_zone_vnc_cession)

			if df.at[ind, "Compe PCO Mise au rebut R et P"] != "" and v_zone_vnc_mar =="X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compe PCO Mise au rebut R et P"], v_sl, 
					find_flux(df.at[ind, "Compe PCO Mise au rebut R et P"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product VNC Mise au rebut R et P"])

			if df.at[ind, "Compe PCO Mise au rebut R et P"] != "" and v_zone_vnc_mar !="X" and v_zone_vnc_mar !="" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compe PCO Mise au rebut R et P"], v_sl, 
					find_flux(df.at[ind, "Compe PCO Mise au rebut R et P"],v_code_flux_bilan,v_code_flux_pnl), v_zone_vnc_mar)

			if df.at[ind, "Compe PCO Produit de cession"] != "" and v_zone_produit_cession == "X":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compe PCO Produit de cession"], v_sl, 
					find_flux(df.at[ind, "Compe PCO Produit de cession"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product Produit cession PS"])

			if df.at[ind, "Compe PCO Produit de cession"] != "" and v_zone_produit_cession != "X" and v_zone_produit_cession != "":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compe PCO Produit de cession"], v_sl, 
					find_flux(df.at[ind, "Compe PCO Produit de cession"],v_code_flux_bilan,v_code_flux_pnl), v_zone_produit_cession)

			if df.at[ind, "Compe PCO Valeur Brute Aux. Impairment"] != "" and v_zone_valeur_impairment == "X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compe PCO Valeur Brute Aux. Impairment"], v_sl, 
					find_flux(df.at[ind, "Compe PCO Valeur Brute Aux. Impairment"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product Valeur Brut Impairment"])

			if df.at[ind, "Compe PCO Valeur Brute Aux. Impairment"] != "" and v_zone_valeur_impairment != "X" and v_zone_valeur_impairment != "":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compe PCO Valeur Brute Aux. Impairment"], v_sl, 
					find_flux(df.at[ind, "Compe PCO Valeur Brute Aux. Impairment"],v_code_flux_bilan,v_code_flux_pnl), v_zone_valeur_impairment)

			if df.at[ind, "Compte PCO dotation Impairment "] != "" and v_zone_dot_impairment == "X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO dotation Impairment "], v_sl, 
					find_flux(df.at[ind, "Compte PCO dotation Impairment "],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product Dotation Impairment"])

			if df.at[ind, "Compte PCO dotation Impairment "] != "" and v_zone_dot_impairment != "X" and v_zone_dot_impairment != "":
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO dotation Impairment "], v_sl, 
					find_flux(df.at[ind, "Compte PCO dotation Impairment "],v_code_flux_bilan,v_code_flux_pnl), v_zone_dot_impairment)

			if df.at[ind, "Compte provision dépréciation actif PCO"] != "" and v_zone_prov_dep == "X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte provision dépréciation actif PCO"], v_sl, 
					find_flux(df.at[ind, "Compte provision dépréciation actif PCO"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "Zone Product Provision dépréciation Actif"])

			if df.at[ind, "Compte provision dépréciation actif PCO"] != "" and v_zone_prov_dep != "X" and v_zone_prov_dep != "" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte provision dépréciation actif PCO"], v_sl, 
					find_flux(df.at[ind, "Compte provision dépréciation actif PCO"],v_code_flux_bilan,v_code_flux_pnl), v_zone_prov_dep)

			if df.at[ind, "Compte PCO dotation provision depreciation actif courant"] != "" and v_zone_dot_prov_actif_courant == "X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO dotation provision depreciation actif courant"], v_sl, 
					find_flux(df.at[ind, "Compte PCO dotation provision depreciation actif courant"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "zone product dot prov actif courrant"])

			if df.at[ind, "Compte PCO dotation provision depreciation actif courant"] != "" and v_zone_dot_prov_actif_courant != "X" and v_zone_dot_prov_actif_courant != "" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO dotation provision depreciation actif courant"], v_sl, 
					find_flux(df.at[ind, "Compte PCO dotation provision depreciation actif courant"],v_code_flux_bilan,v_code_flux_pnl), v_zone_dot_prov_actif_courant)

			if df.at[ind, "Compte PCO reprise provision depreciation actif courant"] != "" and v_zone_reprise_prov_actif_courant == "X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO reprise provision depreciation actif courant"], v_sl, 
					find_flux(df.at[ind, "Compte PCO reprise provision depreciation actif courant"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "zone product rep, prov dep,actif courrant"])

			if df.at[ind, "Compte PCO reprise provision depreciation actif courant"] != "" and v_zone_reprise_prov_actif_courant != "X" and v_zone_reprise_prov_actif_courant != "" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO reprise provision depreciation actif courant"], v_sl, 
					find_flux(df.at[ind, "Compte PCO reprise provision depreciation actif courant"],v_code_flux_bilan,v_code_flux_pnl), v_zone_reprise_prov_actif_courant)

			if df.at[ind, "Compte PCO dotation provision depreciation actif exceptionnel "] != "" and v_zone_prov_dep_excep =="X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO dotation provision depreciation actif exceptionnel "], v_sl, 
					find_flux(df.at[ind, "Compte PCO dotation provision depreciation actif exceptionnel "],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "zone product prov depr actif exceptionnelle"])

			if df.at[ind, "Compte PCO dotation provision depreciation actif exceptionnel "] != "" and v_zone_prov_dep_excep !="X" and v_zone_prov_dep_excep !="" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO dotation provision depreciation actif exceptionnel "], v_sl, 
					find_flux(df.at[ind, "Compte PCO dotation provision depreciation actif exceptionnel "],v_code_flux_bilan,v_code_flux_pnl), v_zone_prov_dep_excep)

			if df.at[ind, "Compte PCO reprise provision depreciation actif exceptionnel"] != "" and v_zone_reprise_prov_dep_excep =="X" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO reprise provision depreciation actif exceptionnel"], v_sl, 
					find_flux(df.at[ind, "Compte PCO reprise provision depreciation actif exceptionnel"],v_code_flux_bilan,v_code_flux_pnl), df.at[ind, "zone product rer, prov, dep, actif exceptionnelle"])

			if df.at[ind, "Compte PCO reprise provision depreciation actif exceptionnel"] != "" and v_zone_reprise_prov_dep_excep !="X" and v_zone_reprise_prov_dep_excep !="" :
				add_code ("", v_code_mouvement, v_category, df.at[ind, "Compte PCO reprise provision depreciation actif exceptionnel"], v_sl, 
					find_flux(df.at[ind, "Compte PCO reprise provision depreciation actif exceptionnel"],v_code_flux_bilan,v_code_flux_pnl), v_zone_reprise_prov_dep_excep)

	if v_action == "Exception" :								# si c'est une exception

		if v_sl == "X" : 										# si le code SL est renseigné

			df2 = df[(df['Category'] == v_sous_category)]		#création d'un dataframe ne contenant que la catégorie égale = sous_category de la matrice
			df2.reset_index(drop=True, inplace=True)			#reset de l'index car lors de la copie on récupère le n° d'index de la category

			for y in df2.index : 

				if df2.at[y, "Compte Valeur Brut Aux. PCO (Drive)"] != "" and v_zone_produit != "":
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte Valeur Brut Aux. PCO (Drive)"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_produit #maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte Valeur Brut Aux. PCO (Drive)"], v_sl, 
							find_flux(df2.at[y, "Compte Valeur Brut Aux. PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_produit)

				if df2.at[y, "Compte Amort. Lin.PCO (Drive)"] != "" and v_zone_amort_lin != "":
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte Amort. Lin.PCO (Drive)"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_amort_lin
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte Amort. Lin.PCO (Drive)"], v_sl, 
							find_flux(df2.at[y, "Compte Amort. Lin.PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_amort_lin)

				if df2.at[y, "Compte Dotation Lin. PCO (Drive)"] != "" and v_zone_dot_lin !="":
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte Dotation Lin. PCO (Drive)"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_dot_lin		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte Dotation Lin. PCO (Drive)"], v_sl, 
							find_flux(df2.at[y, "Compte Dotation Lin. PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_dot_lin)

				if df2.at[y, "Compte Amts dérogatoire PCO (Drive)"] != ""  and v_zone_amort_dero !="":
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte Amts dérogatoire PCO (Drive)"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_amort_dero		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte Amts dérogatoire PCO (Drive)"], v_sl, 
							find_flux(df2.at[y, "Compte Amts dérogatoire PCO (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_amort_dero)

				if df2.at[y, "Compte PCO Dotation Déro. (Drive)"] != "" and v_zone_dot_dero !="":
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte PCO Dotation Déro. (Drive)"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_dot_dero		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte PCO Dotation Déro. (Drive)"], v_sl, 
							find_flux(df2.at[y, "Compte PCO Dotation Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_dot_dero)

				if df2.at[y, "Compte PCO Reprise Déro. (Drive)"] != "" and v_zone_reprise_dero != "" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte PCO Reprise Déro. (Drive)"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_reprise_dero		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte PCO Reprise Déro. (Drive)"], v_sl, 
							find_flux(df2.at[y, "Compte PCO Reprise Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_reprise_dero)

				if df2.at[y, "Compte PCO Reliquat Déro. (Drive)"] != "" and v_zone_reliquat_dero !="":
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte PCO Reliquat Déro. (Drive)"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_reliquat_dero		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte PCO Reliquat Déro. (Drive)"], v_sl, 
							find_flux(df2.at[y, "Compte PCO Reliquat Déro. (Drive)"],v_code_flux_bilan,v_code_flux_pnl), v_zone_reliquat_dero)

				if df2.at[y, "Compe PCO VNC Cessions PS"] != "" and v_zone_vnc_cession != "" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compe PCO VNC Cessions PS"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_vnc_cession		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compe PCO VNC Cessions PS"], v_sl, 
							find_flux(df2.at[y, "Compe PCO VNC Cessions PS"],v_code_flux_bilan,v_code_flux_pnl), v_zone_vnc_cession)

				if df2.at[y, "Compe PCO Mise au rebut R et P"] != "" and v_zone_vnc_mar !="" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compe PCO Mise au rebut R et P"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_vnc_mar		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compe PCO Mise au rebut R et P"], v_sl, 
							find_flux(df2.at[y, "Compe PCO Mise au rebut R et P"],v_code_flux_bilan,v_code_flux_pnl), v_zone_vnc_mar)

				if df2.at[y, "Compe PCO Produit de cession"] != "" and v_zone_produit_cession != "":
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compe PCO Produit de cession"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_produit_cession		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compe PCO Produit de cession"], v_sl, 
							find_flux(df2.at[y, "Compe PCO Produit de cession"],v_code_flux_bilan,v_code_flux_pnl), v_zone_produit_cession)

				if df2.at[y, "Compe PCO Valeur Brute Aux. Impairment"] != "" and v_zone_valeur_impairment != "" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compe PCO Valeur Brute Aux. Impairment"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_valeur_impairment		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compe PCO Valeur Brute Aux. Impairment"], v_sl, 
							find_flux(df2.at[y, "Compe PCO Valeur Brute Aux. Impairment"],v_code_flux_bilan,v_code_flux_pnl), v_zone_valeur_impairment)

				if df2.at[y, "Compte PCO dotation Impairment "] != "" and v_zone_dot_impairment != "" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte PCO dotation Impairment "]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_dot_impairment		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte PCO dotation Impairment "], v_sl, 
							find_flux(df2.at[y, "Compte PCO dotation Impairment "],v_code_flux_bilan,v_code_flux_pnl), v_zone_dot_impairment)

				if df2.at[y, "Compte provision dépréciation actif PCO"] != "" and v_zone_prov_dep != "" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte provision dépréciation actif PCO"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_prov_dep		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte provision dépréciation actif PCO"], v_sl, 
							find_flux(df2.at[y, "Compte provision dépréciation actif PCO"],v_code_flux_bilan,v_code_flux_pnl), v_zone_prov_dep)

				if df2.at[y, "Compte PCO dotation provision depreciation actif courant"] != "" and v_zone_dot_prov_actif_courant != "" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte PCO dotation provision depreciation actif courant"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_dot_prov_actif_courant		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte PCO dotation provision depreciation actif courant"], v_sl, 
							find_flux(df2.at[y, "Compte PCO dotation provision depreciation actif courant"],v_code_flux_bilan,v_code_flux_pnl), v_zone_dot_prov_actif_courant)

				if df2.at[y, "Compte PCO reprise provision depreciation actif courant"] != "" and v_zone_reprise_prov_actif_courant != "" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte PCO reprise provision depreciation actif courant"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_reprise_prov_actif_courant		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte PCO reprise provision depreciation actif courant"], v_sl, 
							find_flux(df2.at[y, "Compte PCO reprise provision depreciation actif courant"],v_code_flux_bilan,v_code_flux_pnl), v_zone_reprise_prov_actif_courant)

				if df2.at[y, "Compte PCO dotation provision depreciation actif exceptionnel "] != "" and v_zone_prov_dep_excep !="" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte PCO dotation provision depreciation actif exceptionnel "]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_prov_dep_excep		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte PCO dotation provision depreciation actif exceptionnel "], v_sl, 
							find_flux(df2.at[y, "Compte PCO dotation provision depreciation actif exceptionnel "],v_code_flux_bilan,v_code_flux_pnl), v_zone_prov_dep_excep)

				if df2.at[y, "Compte PCO reprise provision depreciation actif exceptionnel"] != "" and v_zone_reprise_prov_dep_excep !="" :
					v_temp_cle = str(v_code_mouvement) + "|" + v_sous_category + "|" + str(df2.at[y, "Compte PCO reprise provision depreciation actif exceptionnel"]) + "|" + v_sl 
					number = len(user_exit[user_exit['cle_exception']==v_temp_cle])	#recherche de la clé c-a-d une combinaison hors nature de retraitement
					if number >0:
						user_exit.loc[user_exit["cle_exception"] == v_temp_cle, "Nature_retraitement"] = v_zone_reprise_prov_dep_excep		#maj des valeurs dans user_exit
					else :	
						add_code ("", v_code_mouvement, v_sous_category, df2.at[y, "Compte PCO reprise provision depreciation actif exceptionnel"], v_sl, 
							find_flux(df2.at[y, "Compte PCO reprise provision depreciation actif exceptionnel"],v_code_flux_bilan,v_code_flux_pnl), v_zone_reprise_prov_dep_excep)

		#df2.drop(df2.index,inplace=True)					#Suppression du dataframe temporaire

	if v_action == "Valeur unique" :						# si c'est une exception

		if v_sl == "X" : 									# si le code SL est renseigné
			add_code ("", "", v_sous_category, v_compte, v_sl, find_flux(v_compte,v_code_flux_bilan,v_code_flux_pnl), v_zone_produit)
		else :
			add_code ("", "", "", v_compte, "", find_flux(v_compte,v_code_flux_bilan,v_code_flux_pnl), v_zone_produit)


# check de l'unicité sur le couple Code Mouvement|Categorie|compte|nature et SL 

user_exit.drop(0, axis = 0, inplace=True)					#Suppression de la 1ère ligne qui contient les valeurs (ex string) lors de la création des dataframes

#user_exit['Compte'] = user_exit['Compte'].apply(np.int64)	#changement du n° de compte en int64 pour éviter des 0 au cas ou souhaiterait faire une clé unique
#user_exit.insert (0,"cle",user_exit["Code_mouvement"].astype(str) + "|" + user_exit["Categorie"] + "|"+ user_exit["Compte"].astype(str) + "|"+ user_exit["Nature_retraitement"])						#création de la clé unique Code_societe + Code_id
df5 = user_exit.groupby(['cle'], as_index=False).size()		#création d'un dataframe contenant la clé et le count(*)
print ("Combinaison ayant plus d'une valeur :")
print(df5['cle'][df5['size'] != 1])							#impression des clés en erreur
#user_exit = user_exit[user_exit.Compte != ""]				#Supression des lignes dont la colonne compte est vide
user_exit.drop(columns=['cle'], inplace=True)				#supression de la colonne clé
user_exit = user_exit[['Plan_evaluation', 'S/L_Partenaire', 'Categorie','Code_mouvement', 'Compte', 'Nature_retraitement' ]]
user_exit.insert(loc=4, column='N°_Compte', value="")
user_exit.to_excel(r'C:\Perso\Carrefour\user_exit.xlsx', index=False)
