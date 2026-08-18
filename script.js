// Lien Google Apps Script Web App
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyaI9N3V34TTJ6ipBOC6yF2y4nxWyRXGz4MYeKhogu5MfnTRXMTFiGOCyU4xNIgIEseUA/exec";


// Envoi du formulaire d'inscription
document
.getElementById("inscriptionForm")
.addEventListener("submit", function(e) {

    e.preventDefault();


    // Récupération des informations du formulaire
    const data = {

        // Responsable légal 1
        parent_type:
        document.getElementById("parent_type")?.value || "",

        parent_nom:
        document.getElementById("parent_nom").value,

        parent_prenom:
        document.getElementById("parent_prenom").value,


        // Adresse
        adresse:
        document.getElementById("adresse")?.value || "",

        code_postal:
        document.getElementById("code_postal")?.value || "",

        ville:
        document.getElementById("ville")?.value || "",


        // Contact
        telephone:
        document.getElementById("telephone").value,

        email:
        document.getElementById("email").value,


        // Responsable légal 2
        parent2_type:
        document.getElementById("parent2_type")?.value || "",

        parent2_nom:
        document.getElementById("parent2_nom")?.value || "",

        parent2_prenom:
        document.getElementById("parent2_prenom")?.value || "",

        parent2_telephone:
        document.getElementById("parent2_telephone")?.value || "",

        parent2_email:
        document.getElementById("parent2_email")?.value || "",


        // Enfants (1 à 5)
        nom_1:
        document.getElementById("nom_1")?.value || "",

        prenom_1:
        document.getElementById("prenom_1")?.value || "",

        date_naissance_1:
        document.getElementById("date_naissance_1")?.value || "",

        niveau_demande_1:
        document.getElementById("niveau_demande_1")?.value || "",

        departement_1:
        document.getElementById("departement_1")?.value || "",


        nom_2:
        document.getElementById("nom_2")?.value || "",

        prenom_2:
        document.getElementById("prenom_2")?.value || "",


        nom_3:
        document.getElementById("nom_3")?.value || "",

        prenom_3:
        document.getElementById("prenom_3")?.value || "",


        nom_4:
        document.getElementById("nom_4")?.value || "",

        prenom_4:
        document.getElementById("prenom_4")?.value || "",


        nom_5:
        document.getElementById("nom_5")?.value || "",

        prenom_5:
        document.getElementById("prenom_5")?.value || "",


        // Signature électronique
        signature_image:
        document.getElementById("signature_image")?.value || ""

    };


    // Envoi vers Google Apps Script
    fetch(SCRIPT_URL, {

        method: "POST",

        body: JSON.stringify(data)

    })


    .then(response => response.json())


    .then(result => {

        if(result.status === "ok") {

            alert(
              "Inscription envoyée avec succès !"
            );

            document
            .getElementById("inscriptionForm")
            .reset();

        }

        else {

            alert(
              "Erreur : " + result.message
            );

        }

    })


    .catch(error => {

        console.error(error);

        alert(
          "Erreur lors de l'envoi de l'inscription"
        );

    });


});
