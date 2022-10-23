const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const { DefaultAzureCredential } = require("@azure/identity");

require("isomorphic-fetch");

const credential = new DefaultAzureCredential();
const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ["https://graph.microsoft.com/.default"] }  );

const client = Client.initWithMiddleware({
    defaultVersion: "v1.0",
	debugLogging: true,
	authProvider,
});
    
module.exports = async function (context, req) {

    var userObjectId = (req.query.objectId || (req.body && req.body.objectId));
    var b2cfhirproxySPId = (req.query.clientId || (req.body && req.body.clientId));
    
    console.log("userObjectId: " + userObjectId);
    console.log("b2cfhirproxySPId: " + b2cfhirproxySPId);
    
    let userGroups = [];
    let userAppRoles = [];
    let spObjs = [];
    let jwtRoles = [];
    let jwtGroups = [];

    // retrieve user groups
    await client
        .api("/users/"+userObjectId+"/memberOf/microsoft.graph.group?$count=true&$select=id,displayName")
        .get()
        .then( (res) => {
            userGroups = res.value;
        })

    // build the jwtGroups array
    userGroups.forEach( (group) => {
        console.log("group:"+group.displayName);
        jwtGroups.push(group.displayName);
    });

    // retrieve user assigned roles
    await client
        .api("/users/"+userObjectId+"/appRoleAssignments?$select=appRoleId,resourceId")
        .get()
        .then( (res) => {
            userAppRoles = res.value;
        })

    // decode the app role ids if any  
    if (!userAppRoles.length) 
    {
        console.log("User (objectId:" + userObjectId + ") has no assigned app roles");

    } else {
    
        // retrieve the list of app roles for the given SP
        await client
            .api("/servicePrincipals?$filter=appId eq '"+b2cfhirproxySPId+"' &$select=id,appId,appRoles")
            .get()
            .then((res) => {
                spObjs = res.value;
            })

        // for every sp's id that matches the user app role resourceId, 
        // loop thru the sp's appRoles to find a match to the user's app role id 
        // to retrieve the sp's appRole value.
        spObjs.forEach( (spObj) => {
            userAppRoles.forEach( (userAppRole) => {
                if (spObj.id == userAppRole.resourceId) {
                    spObj.appRoles.forEach( (appRole) => {
                        if (appRole.id == userAppRole.appRoleId){
                            jwtRoles.push(appRole.value);
                        }
                    });
                } 
            });
        });    
    }

    jwtRoles.forEach( (r) => { console.log("role:"+r)} );

    context.res = {
        body: {"roles": jwtRoles, "groups": jwtGroups},
        contenttype: "application/json" 
    }
}
