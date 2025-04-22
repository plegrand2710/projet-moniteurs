(function() {
    // Fonction pour formater une date au format "YYYY-MM-DD HH:MM:SS"
    function formatDate(date) {
      var yyyy = date.getFullYear();
      var mm = String(date.getMonth() + 1).padStart(2, '0'); // Mois : 01 à 12
      var dd = String(date.getDate()).padStart(2, '0');
      var hh = String(date.getHours()).padStart(2, '0');
      var mi = String(date.getMinutes()).padStart(2, '0');
      var ss = String(date.getSeconds()).padStart(2, '0');
      return yyyy + '-' + mm + '-' + dd + ' ' + hh + ':' + mi + ':' + ss;
    }
  
    function formatShortDate(date) {
      var mm = String(date.getMonth() + 1).padStart(2, '0'); // Mois : 01 à 12
      var dd = String(date.getDate()).padStart(2, '0');
      return dd + '-' + mm;
    }

    // Vérifier que la variable 'list' est définie et contient des événements
    if (typeof list === "undefined" || !list._array || list._array.length === 0) {
      console.error("La variable 'list' n'est pas définie ou ne contient aucun événement.");
      return;
    }
  
    // Initialiser le contenu CSV avec l'en-tête (colonnes séparées par des virgules)
    var csvContent = "Début,Fin,Lieu,Nom\n";
  
    // Parcourir chaque événement et ajouter une ligne au CSV
    list._array.forEach(function(event) {
      console.log(event);

      var debut = event.startDate ? formatDate(event.startDate) : "";
      var fin = event.endDate ? formatDate(event.endDate) : "";
      var lieu = event.location || "";
      var nom = event.name || "";
      
      // Chaque champ est encadré par des guillemets, séparé par des virgules
      csvContent += '"' + debut + '","' + fin + '","' + lieu + '","' + nom + '"\n';
    });
  
    var firstDate = list._array[0];
    console.log(firstDate);

    var dateString = firstDate.startDate;
    dateString = formatShortDate(new Date(dateString));
    var fileName = "Evenements_Travail_Semaine_Prochaine_" + dateString + ".csv";

    // Création d'un Blob à partir du contenu CSV
    var blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    var link = document.createElement("a");
    var url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", fileName);
    link.style.visibility = 'hidden';
    
    // Ajouter le lien au document, simuler le clic puis le retirer
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    console.log("Fichier CSV généré et téléchargement lancé.");
  })();
  