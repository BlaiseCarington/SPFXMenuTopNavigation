//import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';


//import { Dialog } from '@microsoft/sp-dialog';

//import * as strings from 'MenuByBibliothequeApplicationCustomizerStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
interface IListItem {
  Title: string;
  DefaultViewUrl: string;
}
interface IListResponse {
  value: IListItem[];
}

//import styles from './MenuByBibliotheque.module.scss';
//const LOG_SOURCE: string = 'MenuByBibliothequeApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMenuByBibliothequeApplicationCustomizerProperties {
  // This is an example; replace with your own property
  //testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class MenuByBibliothequeApplicationCustomizer
  extends BaseApplicationCustomizer<IMenuByBibliothequeApplicationCustomizerProperties> {
  private _topPlaceholder?: PlaceholderContent;
  public onInit(): Promise<void> {
    this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top,
      { onDispose: this._onDispose }
    );

    if (!this._topPlaceholder) {
      console.warn('Placeholder Top non disponible.');
      return Promise.resolve();
    }

    // Inject CSS directly
    const styleTag = document.createElement('style');
    styleTag.innerHTML = `
    * {
      margin: 0px;
      padding: 0px;
      font-family: Avenir, sans-serif;
    }

    nav {
      width: 100%;
      margin: 0 auto;
      background-color: white;
      position: sticky;
      z-index: 100000; /* Doit être plus grand que celui du header SharePoint */
      top: 0px;
    }

    nav ul {
      list-style-type: none;
    }

    nav ul li {
      float: left;
      width: auto;
      text-align: center;
      position: relative;
      font-size: 7pt;
      margin-right: 3px;
    }

    nav ul li:last-child {
      margin-right: 0;
    }

    nav ul::after {
      content: "";
      display: table;
      clear: both;
    }

    nav a {
      display: block;
      text-decoration: none;
      color: black;
      border-bottom: 2px solid transparent;
      padding: 10px 0px;
      font-size: 10pt;        /* ✅ Augmente la taille */
      font-weight: bold;      /* ✅ Met en gras */
    }


    nav a:hover {
      color: black;
      border-bottom: 2px solid gold;
    }

    .sous, .soussous {
      display: none;
      box-shadow: 0px 1px 2px #CCC;
      background-color: white;
      position: absolute;
      z-index: 2147483647; /* Maximum autorisé dans les navigateurs  permet de mettre la navigation au dessus de tous les composants de la page sharepoint*/
    }

    .sous {
      width: fit-content;
      white-space: nowrap;
      padding: 5px 0; /* optionnel : pour aérer un peu */
    }


    nav > ul li:hover .sous {
      display: block;
    }

    .sous li, .soussous li {
      float: none;
      width: 100%;
      text-align: left;
    }

    .sous a, .soussous a {
      padding: 10px;
      border-bottom: none;
      font-size: 11pt;      /* ✅ Taille de police augmentée */
      font-weight: bold;    /* ✅ Texte en gras (optionnel) */
    }


    .sous a:hover, .soussous a:hover {
      background-color: rgba(200, 200, 200, 0.1);
    }

    .deroulant > a::after {
      content: "";
      font-size: 12px;
    }

    .deroulant:has(> .sous) > a::after {
      content: "▼";
      font-size: 12px;
    }

    nav ul li ul li {
      float: left;
      width: 25%;
      text-align: center;
      position: relative;
    }

    nav > ul li ul li:hover .soussous {
      display: block;
    }
  
    nav ul li.deroulant:nth-child(1) { background-color: #DAF7A6; }
    nav ul li.deroulant:nth-child(2) { background-color: #6794E0; }
    nav ul li.deroulant:nth-child(3) { background-color: #DD69B2; }
    nav ul li.deroulant:nth-child(4) { background-color: #FD936B; }
    nav ul li.deroulant:nth-child(5) { background-color: #9D6BC2; }
    nav ul li.deroulant:nth-child(6) { background-color: #68BF8E; }
    nav ul li.deroulant:nth-child(7) { background-color: #CCCCCC; }
    nav ul li.deroulant:nth-child(8) { background-color: #EEB4D9; }
    nav ul li.deroulant:nth-child(9) { background-color: #FEC9B5; }
    nav ul li.deroulant:nth-child(10) { background-color: #CEB5E0; }
    nav ul li.deroulant:nth-child(11) { background-color: #B2B2B2; }
    nav ul li.deroulant:nth-child(12) { background-color: #FFC300; }
  `;
    document.head.appendChild(styleTag);

    // Inject your menu HTML
    this._topPlaceholder.domElement.innerHTML = `
    <nav>
      <ul>
        <li class="deroulant" id="Accueil"><a href="${this.context.pageContext.web.absoluteUrl}">&ensp;Accueil &ensp;</a></li>
        <li class="deroulant" id="A.Avant-projet"><a href="#">&ensp;A.Avant-projet &ensp;</a></li>
        <li class="deroulant" id="B.Initiation"><a href="#">&ensp;B.Initiation &ensp;</a></li>
        <li class="deroulant" id="C.Planification"><a href="#">&ensp;C.Planification &ensp;</a></li>
        <li class="deroulant" id="D.Execution"><a href="#">&ensp;D.Execution &ensp;</a></li>
        <li class="deroulant" id="E.Fermeture"><a href="#">&ensp;E.Fermeture &ensp;</a></li>
        <li class="deroulant" id="F.Gestion_continue"><a href="#">&ensp;F.Gestion_continue &ensp;</a></li>
        <li class="deroulant" id="G.ExterneUL_Spro-INI"><a href="#">&ensp;G.ExterneUL_Spro-INI &ensp;</a></li>
        <li class="deroulant" id="H.ExterneUL_Spro-PLA"><a href="#">&ensp;H.ExterneUL_Spro-PLA &ensp;</a></li>
        <li class="deroulant" id="I.ExterneUL_EXE"><a href="#">&ensp;I.ExterneUL_EXE &ensp;</a></li>
        <li class="deroulant" id="J.Utilitaires"><a href="#">&ensp;J.Utilitaires &ensp;</a></li>
        <li class="deroulant" id="Administration"><a href="#">&ensp;Administration &ensp;</a></li>
      </ul>
    </nav>`;
    this._loadAAvantprojet();
    this._loadBInitiation();
    this._loadCPlanification();
    this._loadDExecution();
    this._loadEFermeture();
    this._loadFGestion_continue();
    this._loadGExterneULSproINI();
    this._loadHExterneULSproPLA();
    this._loadIExterneULENT();
    this._loadJUtilitaires();
    this._loadAdmin();
    return Promise.resolve();
  }

  private _loadAAvantprojet(): void {
    const internalNames = ['A01Besoin', 'A02Expertise'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "A.Avant-projet");
  }
  private _loadBInitiation(): void {
    const internalNames = ['B01PFT', 'B02Plan_Devis'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "B.Initiation");
  }

  private _loadCPlanification(): void {
    const internalNames = ['C01Terrain', 'C02Plan_Devis', 'C03Service'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "C.Planification");
  }

  private _loadDExecution(): void {
    const internalNames = ['D01Permis_Services_publics', 'D02Contrat_entrepreneur', 'D03Surveillance', 'D04QRT', 'D05Mise_en_service'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "D.Execution");
  }

  private _loadEFermeture(): void {
    const internalNames = ['E01Lecon_apprise', 'E02Bilan'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "E.Fermeture");
  }

  private _loadFGestion_continue(): void {
    const internalNames = ['F01Gouvernance', 'F02Partie_prenante', 'F03Intervenant', 'F04Charte_MOP', 'F04ADossier_opportunite', 'F04BDossier_affaires', 'F05Gestion_contractuelle', 'F06Gestion_budgetaire', 'F07Gestion_administrative', 'F08Echeancier', 'F09Analyse_valeur', 'F10Risque', 'F11Audit', 'F12Developpement_durable', 'F13Œuvre_art', 'F14Communication', 'F15BIM', 'F16Photo'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "F.Gestion_continue");
  }

  private _loadGExterneULSproINI(): void {
    const internalNames = ['G01PFT', 'G02Plan_Devis', 'G14Communication'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "G.ExterneUL_Spro-INI");
  }

  private _loadHExterneULSproPLA(): void {
    const internalNames = ['H02Plan_Devis', 'H14Communication'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "H.ExterneUL_Spro-PLA");
  }

  private _loadIExterneULENT(): void {
    const internalNames = ['I02Contrat_entrepreneur', 'I03Surveillance', 'I04QRT', 'I05Mise_en_service', 'I14Communication'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "I.ExterneUL_EXE");
  }

  private _loadJUtilitaires(): void {
    const internalNames = ['J03Intervenant', 'J08Echeancier', 'J16Photo'];  //  noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "J.Utilitaires");
  }

  private _loadAdmin(): void {
    const internalNames = ['Gestiondesexternesparprojet'];  // noms logiques
    const filters = internalNames
      .map(name => `RootFolder/Name eq '${name}'`)
      .join(' or ');
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$filter=Hidden eq false and (${filters})` +
      `&$select=Title,DefaultViewUrl`;
    this._generationSousMenu(requestUrl, "Administration");
  }


  private _generationSousMenu(requete: string, IdMenu: string): void {
    let NbFichier: number = 0;
    this.context.spHttpClient.get(requete, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: IListResponse) => {
        const lists: IListItem[] = data.value;
        const Menu = document.getElementById(IdMenu);

        if (Menu === null) {
          console.log('L\'élément est null');
        } else {
          const ul = document.createElement('ul');
          ul.className = "sous";
          lists.forEach(listItem => {
            const li = document.createElement('li');
            const a = document.createElement('a');
            //ul.className = styles.sous;
            a.href = listItem.DefaultViewUrl;
            //a.target = '_blank';
            a.rel = 'noopener noreferrer';

            a.style.textDecoration = 'none';
            a.style.fontSize = '1.1em';
            a.textContent = `${listItem.Title}`;

            if (IdMenu !== "Administration") {
              //Ajout du nombre de fichiers
              this.countAllFilesInLibrary(`${listItem.Title}`)
                .then(count => {
                  //a.textContent = `${listItem.Title} (${count})`;
                  //  alert(`${count}`);
                  NbFichier += count;
                  const link = Menu.querySelector('a');
                  if (link !== null) {
                    if (NbFichier !== 0) {
                      link.textContent = `${IdMenu} (${NbFichier})`;
                    }
                  }
                  //Menu.textContent = `${IdMenu} (${NbFichier})`;
                })
                .catch(err => console.error('Échec récupération count :', err));
            }else{a.textContent = `Gestion-accès`;}




            li.appendChild(a);
            ul.appendChild(li);
          });

          Menu.appendChild(ul);
        }
      })
      .catch(error => {
        console.error('Erreur de chargement des bibliothèques :', error);
      });
  }

  /* private _getFileCount(libraryTitle: string): Promise<number> {

    const webUrl = this.context.pageContext.web.absoluteUrl;
    const endpoint = `${webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(libraryTitle)}')?$select=ItemCount`;
    // const endpoint = `${webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(libraryTitle)}')/items?$filter=FSObjType eq 0&$top=1&$count=true`;
    return this.context.spHttpClient
      .get(endpoint, SPHttpClient.configurations.v1)
      .then(res => {
        if (!res.ok) {
          return Promise.reject(new Error(`HTTP ${res.status} – ${res.statusText}`));
        }
        return res.json();
      })
      .then((data: any) => {
        return data.ItemCount as number;

      });


  } */
  
  private async countAllFilesInLibrary(libraryTitle: string): Promise<number> {
    // Récupérer l'URL du dossier racine de la bibliothèque
    const rootFolderUrl: string = await this.getLibraryRootFolderUrl(libraryTitle);
    // Appeler la fonction récursive pour compter
    const total: number = await this.countFilesRecursively(rootFolderUrl);
    return total;
  }
  
  private async getLibraryRootFolderUrl(libraryTitle: string): Promise<string> {
    const endpoint: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(libraryTitle)}')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Impossible de récupérer le dossier racine pour la bibliothèque "${libraryTitle}". Vérifiez le nom.`);
    }
    const data = await response.json();
    return data.RootFolder.ServerRelativeUrl;
  }

  private async countFilesRecursively(folderServerRelativeUrl: string): Promise<number> {
    let count = 0;

    // 1) Récupérer tous les fichiers du dossier courant
    const filesEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/Files?$select=Name`;
    const filesResponse: SPHttpClientResponse = await this.context.spHttpClient.get(filesEndpoint, SPHttpClient.configurations.v1);
    if (!filesResponse.ok) {
      throw new Error(`Erreur lors de la récupération des fichiers du dossier "${folderServerRelativeUrl}".`);
    }
    const filesJson = await filesResponse.json();
    const filesArray = filesJson.value as any[];
    count += filesArray.length;

    // 2) Récupérer tous les sous-dossiers du dossier courant
    const foldersEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/Folders?$select=Name,ServerRelativeUrl`;
    const foldersResponse: SPHttpClientResponse = await this.context.spHttpClient.get(foldersEndpoint, SPHttpClient.configurations.v1);
    if (!foldersResponse.ok) {
      throw new Error(`Erreur lors de la récupération des sous-dossiers de "${folderServerRelativeUrl}".`);
    }
    const foldersJson = await foldersResponse.json();
    const foldersArray = foldersJson.value as { Name: string; ServerRelativeUrl: string }[];

    // 3) Pour chaque sous-dossier, rappeler la fonction de façon récursive
    for (const folder of foldersArray) {
      // Ignorer le dossier "Forms" ou tout dossier système commun (on peut ajouter un filtre si besoin)
      if (folder.Name === 'Forms') {
        continue;
      }
      const subCount = await this.countFilesRecursively(folder.ServerRelativeUrl);
      count += subCount;
    }

    return count;
  }

  private _onDispose(): void {
    // Nettoyage si nécessaire
  }
}
