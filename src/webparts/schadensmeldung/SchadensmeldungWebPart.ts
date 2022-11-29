import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './SchadensmeldungWebPart.module.scss';
import * as strings from 'SchadensmeldungWebPartStrings';

export interface ISchadensmeldungWebPartProps {
  list: string;
}

export interface ListItem {
  Title: string;
  Id: number;
}

//HTML für einblendbare Teile des Webparts

const unfallGegnerKontaktdatenHtml: string = `
<p class="${styles.paragraph}">
  <span class="${styles.leftItem}">Unfallgegner Kontaktdaten</span> 
  <span class="${styles.rigthItem}">
    <input id="contact" value=""></input>
  </span>
</p>
`;

const verursacherHtml: string = `
<p class="${styles.paragraph}">
  <span class="${styles.leftItem}">Verursacher / beteiligte Person intern</span> 
  <span class="${styles.rigthItem}">
    <select id="responsible">
      <option value="TestA">TestA</option>
      <option value="TestB">TestB</option>
    </select>
  </span>
</p>
`;

const sachgegenstandHtml: string = `
    <p class="${styles.paragraph}">
      <span class="${styles.leftItem}">Sachgegenstand</span> 
      <span class="${styles.rigthItem}">
        <input id="objectDamage" value=""></input>
      </span>
</p>
`;

const fahrzeugSchadenHtml = `
      <p class="${styles.paragraph}">
        <span class="${styles.leftItem}">Eigenes Fahrzeug/Maschine</span> 
        <span class="${styles.rigthItem}">
          <select id="vehicleDamage">
            <option value="Test1">Test1</option>
            <option value="Test2">Test2</option>
          </select>
        </span>
</p>
`;

export default class SchadensmeldungWebPart extends BaseClientSideWebPart<ISchadensmeldungWebPartProps> {


  //Beschr: Methode zu rendern des Webparts
  public render(): void {
    this.domElement.innerHTML = `
    <div>
      <div align="center">
      <p class="${styles.paragraph}">
        <span class="${styles.leftItem}">Datum</span>
        <span class="${styles.rigthItem}">
          <input type="date" id="dateInput" value="">
        </span>
      </p> 
      <p class="${styles.paragraph}">
      <span class="${styles.leftItem}">Art des Schadens</span>
      </p>
      <p >
        <span class="${styles.leftItem}">
          <input type="radio" name="schaden" value="Fahrzeugschaden" id="schaden1">Fahrzeug/Maschine</input>
        </span>
        <span class="${styles.rigthItem}">
          <input type="radio" name="schaden" value="Sachschaden" id="schaden2">Sachschaden</input>
        </span>         
      </p>
      <div class="${styles.paragraph}" id="schadensArt"></div>
      <p class="${styles.paragraph}">
        <span class="${styles.leftItem}">
          <input type="radio" name="geschädigter" value="intern Geschädigter" id="geschädigter1">intern Geschädigter</input>
        </span> 
        <span class="${styles.rigthItem}">
          <input type="radio" name="geschädigter" value="extern Geschädigter" id="geschädigter2">extern Geschädigter</input>
        </span> 
      </p>
      <div class="${styles.paragraph}" id="geschädigterArt"></div>
      <p class="${styles.paragraph}"> 
        <span class="${styles.leftItem}">Hergang des Schadens</span> 
        <span class="${styles.rigthItem}">
          <input id="originOfDamage" value=""></input>
        </span>
      </p>
      <p class="${styles.paragraph}">  
        <span class="${styles.leftItem}">Bilder</span> 
        <span class="${styles.rigthItem}">
          <label for="input" class="${styles['custom-file-upload']}">Upload</label> 
          <input multiple type="file" id="input" accept="image/*" class="${styles.hide}"> 
        </span>
      </p>
      <p class="${styles.paragraph}">
        <button class="${styles['custom-file-upload']}" id="meldung" align="center">Meldung abschicken</button>
      </p>
      <p class="${styles.paragraph}" id="error"></p>
    </div>
    `;
    this.setEventHandlers();
  }

  //Beschr: Methode zur Initalisierung von den eventHandlers
  private setEventHandlers(): void {
    document.getElementById("schaden1").addEventListener('click', () => { this.showDamage(fahrzeugSchadenHtml); });
    document.getElementById("schaden2").addEventListener('click', () => { this.showDamage(sachgegenstandHtml); });
    document.getElementById("geschädigter1").addEventListener('click', () => { this.showDamaged(verursacherHtml); });
    document.getElementById("geschädigter2").addEventListener('click', () => { this.showDamaged(unfallGegnerKontaktdatenHtml + verursacherHtml); });
    document.getElementById("meldung").addEventListener('click', () => { this.uploadReport() })
  }

  //Beschr: Methode zu rendern von Webpartteil der von radio buttons abhängt
  //Input: html sind die HTML strings vor der Klasse
  //TODO
  //Zusammenfassen?
  private showDamage(html: string): void {
    const listContainer: Element = this.domElement.querySelector('#schadensArt');
    listContainer.innerHTML = html;
    return;
  }

  private showDamaged(html: string): void {
    const listContainer: Element = this.domElement.querySelector('#geschädigterArt');
    listContainer.innerHTML = html;
    return;
  }


  //Beschr: Holt sich ausgefüllten Werte des Webparts und lädt sie in eine Liste hoch
  private uploadReport(): void {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.list}')/items`
    const date = (<HTMLInputElement>document.getElementById("dateInput")).value;
    const damageTye: string = this.getDamageType();
    let vehicleDamage: string = "";
    let objectDamage: string = "";
    if (damageTye === (<HTMLInputElement>document.getElementById("schaden1")).value) {
      vehicleDamage = (<HTMLInputElement>document.getElementById("vehicleDamage")).value;
    }
    if (damageTye === (<HTMLInputElement>document.getElementById("schaden2")).value) {
      objectDamage = (<HTMLInputElement>document.getElementById("objectDamage")).value;
    }
    const damagedType: string = this.getDamagedType();
    let contact: string = "";
    let responsible: string = "";
    if (damagedType === (<HTMLInputElement>document.getElementById("geschädigter1")).value) {
      responsible = (<HTMLInputElement>document.getElementById("responsible")).value;
    }
    if (damagedType === (<HTMLInputElement>document.getElementById("geschädigter2")).value) {
      contact = (<HTMLInputElement>document.getElementById("contact")).value;
      responsible = (<HTMLInputElement>document.getElementById("responsible")).value;
    }
    const originOfDamage: string = (<HTMLInputElement>document.getElementById("originOfDamage")).value;

    if (date === "" || damageTye === "" || damagedType === "") {
      this.showError();
      return;
    }
    if (damageTye === (<HTMLInputElement>document.getElementById("schaden2")).value && objectDamage === "") {
      this.showError();
      return;
    }
    if (damagedType === (<HTMLInputElement>document.getElementById("geschädigter2")).value && (responsible === "" || contact === "")) {
      this.showError();
      return;
    }

    if (originOfDamage === "") {
      this.showError();
      return;
    }

    const body: string = JSON.stringify({
      'Title': `Meldung0`,
      'Datum': `${date}`,
      'ArtdesSchadens': `${damageTye}`,
      'Gesch_x00e4_digter': `${damagedType}`,
      'HergangdesSchadens': `${originOfDamage}`,
      'eigenesFahrzeug_x002f_Maschine': `${vehicleDamage}`,
      'Sachschaden': `${objectDamage}`,
      'UnfallgegnerKontaktdaten': `${contact}`,
      'Verursacher': `${responsible}`,
    });

    this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      }
    ).then((response: SPHttpClientResponse): Promise<ListItem> => {
      const error = document.getElementById("error");
      error.innerHTML = "<span style='color: green;'>File submitted</span>";
      return response.json();
    }).then((listItem: ListItem) => {
      this.addAttachament(listItem.Id, 0);
      return;
    })
      .catch((response: SPHttpClientResponse) => {
        console.log(response.json);
        throw Error('Creating Item Failed');
      });
  }
  //Input: itemId ist die ID des Listenelemts an dem die Attachments angehängt werden soll
  //       index ist die Stelle an dem in der Liste der Files geschaute werden soll um die Bilder nacheinander hochzuladen
  //Beschr: Rekursive Methode, die dem Listenelement mit itemID das Bild an Stelle index 
  //        der Fileliste ans Attachement hochlädt. 
  private addAttachament(itemId: number, index: number): void {
    const file = (<HTMLInputElement>document.getElementById("input")).files[index];
    console.log((<HTMLInputElement>document.getElementById("input")).files);
    if (file) {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.list}')/items('${itemId}')/AttachmentFiles/add(FileName='` + file.name + `')`;
      this.context.spHttpClient.post(

        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json",
            "Content-Type": "application/json"
          },
          body: file
        })
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .then(() => {
          this.addAttachament(itemId, index + 1);
        })
        .catch(() => {
          console.log("Adding image failed")
          return;
        });
        
    }
  }

  //Beschr: Diese Methode wird aufgerufen, wenn nicht alle notwendigen Felder des Webparts ausgefüllt sind.
  //        Es wird dann eine Fehlernachricht der dom hinzugefügt.
  private showError(): void {
    const error = document.getElementById("error");
    error.innerHTML = "<span style='color: red;'>Fill out the entire sheet or else..</span>"
  }

  //Beschr: Gibt den ausgewälten Wert der Radiobuttons zurück
  //TODO 
  //Zusammenfassen
  private getDamageType(): string {
    const damageType1 = (<HTMLInputElement>document.getElementById("schaden1"));
    const damageType2 = (<HTMLInputElement>document.getElementById("schaden2"));

    if (damageType1.checked) {
      return damageType1.value;
    } else if (damageType2.checked) {
      return damageType2.value;
    } else {
      return "";
    }
  }

  private getDamagedType(): string {
    const damagedType1 = (<HTMLInputElement>document.getElementById("geschädigter1"));
    const damagedType2 = (<HTMLInputElement>document.getElementById("geschädigter2"));

    if (damagedType1.checked) {
      return damagedType1.value;
    } else if (damagedType2.checked) {
      return damagedType2.value;
    } else {
      return "";
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //Input: value ist der Wert des Textfeldes der Webpartbeschreibung
  //Beschr: Prüft die Beschreibung des Webparts auf Korrektheit
  //TODO
  //Anpassen und prüfen auf existierende Liste
  private validateDescription(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return 'Provide a description';
    }

    return '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('list', {
                  label: strings.ListFieldLabel,
                  onGetErrorMessage: this.validateDescription.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}