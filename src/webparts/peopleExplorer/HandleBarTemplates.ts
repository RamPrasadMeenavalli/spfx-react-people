export const HandleBarTemplates = {
    getTemplate: (type: string): string => {
        let _template: string = "";

        switch (type) {
            case "detailed":
                _template = `<div class={{styles.stackVertical}}>
                <div class='person-card'>
                
                <img src='{{imageUrl}}' class={{styles.userImage}} height='75px' width='75px' /> 
                
                <div class='displayName'><b>{{DisplayName}}</b></div> 
                
                {{#if Title}}
                <div class={{styles.infoLine}}>{{Title}}</div> 
                {{/if}}
                
                {{#if Email}}
                <div class={{styles.infoLine}}>
                <a class='email' href='mailto:{{Email}}'>
                {{Email}}
                </a></div> 
                {{/if}}
                
                {{#if Department}}
                <div class={{styles.infoLine}}>Department : {{Department}}</div> 
                {{/if}}
                
                
                {{#if Office}}
                <div class={{styles.infoLine}}>Office : {{Office}}</div> 
                {{/if}}
                
                {{#if WorkPhone}}
                <div class={{styles.infoLine}}>Work Phone : {{WorkPhone}}</div> 
                {{/if}}
                
                
                {{#if AboutMe}}
                <div class='bio'>
                {{{AboutMe}}}
                </div>
                {{/if}}
                
                </div><!-- end person-card -->
                </div><!-- end vertical stack -->
                
                <style>
                .person-card {
                padding-bottom: 20px;
                }
                .person-card > div {
                margin-top: 5px;
                }
                .displayName {
                margin-bottom: 10px;
                }
                .email {
                color: black !important;
                }
                .bio {
                font-size: 14px;
                font-style: italic;
                }
                </style>
                `
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
                        <div><b>{{DisplayName}}</b></div>
                        <div class={{styles.infoLine}}>{{Title}}</div>
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