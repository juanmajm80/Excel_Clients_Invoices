# Excel_Clients_Invoices

Excel to management clients, benefits and invoices for a very small enterprise or autonomous

This project have been create to help a french friend so all the fileds are in french. My friend is ergotherapist and needed something very simple to manage her clients and the appointment with them and she could control her benefits and her outstanding pacients. She had an Excel file to do this function so I decided mantein Excel but introducing some automatization (macros and formulas) to facilitate her the work.

The new Excel file have 5 parts:

- Part 1 - This is the first sheet and it's only to access to the diferent sheets like a link

- Part 2 - It is formed by 2 sheets and they are two databases:
    - Pacients
       - Have the diferent clients
       - One line is a unique client
       - The fields are divided in 3 blocks:
           - White fields
                - These fields are the cells on which the information collected in the Formulaire Patient will be copied.
                - These fields are:
                    - PRENOM
                    - NOM1
                    - DATE NAISSANCE
                    - PATHOLOGIE
                    - OBSERVATIONS
                    - TELEPHONE1
                    - TELEPHONE2
                    - EMAIL1
                    - EMAIL2
                    - DEPARTEMENT
                    - COMMUNE
                    - ADRESSE
           - Blue fields
                - These fields are cells with formulas and they should not be modified
                - These fields are:
                    - PACIENT - Concatenation of the patient's name and surname
                    - CODE PACIENT - Unique code for each patient
                    - DATE DE SORTIE - Patient creation date
                    - AGE - Age of the patient
                    - AN - Years to the patient's date of birth
                    - MOIS - Meses hasta la fecha de nacimiento del paciente
                    - Nº SEANCES - Months to the patient's date of birth
           - Hidden fields
                - These fields are child cells for calculations and background tasks
                - These fields are:
                    -  DEPARTEMENT+COMMUNE2
                    -  Columna1 - Separation cell
                    -  Columna2 - Separation cell
                    -  PRENOM2 - Receive the information of the Formulaire Pacient cells
                    -  NOM12 - Receive the information of the Formulaire Pacient cells
                    -  DATE NAISSANCE (DD/MM/YYYY) - Receive the information from the Formulaire Pacient cells
                    -  PATHOLOGIE2 - Receive the information of the Formulaire Pacient cells
                    -  OBSERVATIONS3 - Receive the information of the Formulaire Pacient cells
                    -  TELEPHONE12 - Receive the information of the Formulaire Pacient cells
                    -  TELEPHONE23 - Receive the information of the Formulaire Pacient cells
                    -  EMAIL14 - Receive the information of the Formulaire Pacient cells
                    -  EMAIL25 - Receive the information of the Formulaire Pacient cells
                    -  DEPARTMENT - Receive the information of the Formulaire Pacient cells
                    -  COMUNA - Receive the information of the Formulaire Pacient cells
                    -  ADRESSE6 - Receive the information of the Formulaire Pacient cells
                    -  AN2 - Used to calculate the formulas of the Rapport Patients
                    -  NOMBRE PACIENT - Used to calculate the formulas of the Rapport Patients sheet
                    -  no paye - Used to calculate the formulas of the Rapport Patients sheet
                    -  paye - Used to calculate the formulas of the Rapport Patients sheet
    - Seances
       - Have the differents appoinments with the differents clients
       - It can have many lines for the same pacient
       - The fields we wanted for further analysis are:
            - cc
            - z
- Part 3 - It is formed by 3 sheets and they are the formularies to entry the information to the differents fields of the databases sheets:
    - Formulaire Pacient
    - Formulaire Seance
    - Formulaire Facture

- Part 4 - This is a sheet to see individualy the relevant information of each pacient

- Part 5 - It is formed by 3 sheets and they are dashboards to visualizaice the important data of the 2 databases sheets
