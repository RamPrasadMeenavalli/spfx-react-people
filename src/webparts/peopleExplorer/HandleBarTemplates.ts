export const HandleBarTemplates = {
    getTemplate: (type: string): string => {
        let _template: string = "";

        switch (type) {
            case "detailed":
                _template = `<div class={{styles.stackVertical}}>
                    <img 
                        src='{{imageUrl}}' class={{styles.userImage}} 
                        style="margin-bottom:1rem; margin-top:2rem"
                        height="75px"
                        width="75px"
                        />

                        <div><b>{{displayName}}</b></div>
                        <div class={{styles.infoLine}}>{{jobTitle}}</div>
                        <div class={{styles.infoLine}}>{{mail}}</div>

                        <div class={{styles.infoLine}}>Department : {{department}}</div>
                        <div class={{styles.infoLine}}>Office : {{officeLocation}}</div>
                        {{#each phones}}
                            <div class={{../styles.infoLine}}>{{this.type}} : {{this.number}}</div>
                        {{/each}}
                </div>`
                break;
            case "simple":
            default:
                _template = `<div class={{styles.stackHorizontal}}>
                <img 
                    src='{{imageUrl}}' class={{styles.userImage}} 
                    style="margin-right:1rem;"
                    height="50px"
                    width="50px"
                    />
                    <div class={{styles.stackVertical}}>
                        <div><b>{{displayName}}</b></div>
                        <div class={{styles.infoLine}}>{{jobTitle}}</div>
                    </div>
                    
                </div>`;
                break;
        }

        return _template;
    }
}

const sampleObject = { 
    "imageInitials": "AV", 
    "imageUrl": "https://contoso.sharepoint.com/sites/Demos/_layouts/15/userphoto.aspx?accountname=AdeleV%40contoso.onmicrosoft.com&size=M",
    "mail": "AdeleV@contoso.onmicrosoft.com", 
    "id": "XXXXXXXXXXXXXXXXXXXXXXXXXXx", 
    "displayName": "Adele Vance", 
    "givenName": "Adele", 
    "surname": "Vance", 
    "birthday": null, 
    "personNotes": null, 
    "isFavorite": false, 
    "jobTitle": "Retail Manager", 
    "companyName": null, 
    "yomiCompany": null, 
    "department": "Retail", 
    "officeLocation": "18/2111", 
    "profession": null, 
    "userPrincipalName": "AdeleV@contoso.onmicrosoft.com", 
    "imAddress": "sip:adelev@contoso.onmicrosoft.com", 
    "scoredEmailAddresses": [
        { "address": "AdeleV@contoso.onmicrosoft.com", "relevanceScore": -7, "selectionLikelihood": "notSpecified" }], 
    "phones": [{ "type": "business", "number": "+1 425 555 0109" }], 
    "personType": { "class": "Person", "subclass": "OrganizationUser" } 
}