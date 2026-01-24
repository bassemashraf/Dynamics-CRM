/**
 * On Load handler for Work Order form
 * Checks msdyn_servicerequest field and handles account/contact logic
 */
async function onLoadHandler(executionContext) {
    try {
        const formContext = executionContext.getFormContext();
        
        // Get the service request field (Case lookup)
        const serviceRequestField = formContext.getAttribute("msdyn_servicerequest");
        
        if (!serviceRequestField) {
            console.log("Service request field not found");
            return;
        }
        
        const serviceRequestValue = serviceRequestField.getValue();
        
        // If no service request is selected, exit
        if (!serviceRequestValue || serviceRequestValue.length === 0) {
            console.log("No service request selected");
            return;
        }
        
        const caseId = serviceRequestValue[0].id.replace(/[{}]/g, ''); // Remove curly braces
        const caseName = serviceRequestValue[0].name;
        
        console.log("Service Request found:", caseId, caseName);
        
        // Retrieve the Case record to check the Customer field
        const caseRecord = await Xrm.WebApi.retrieveRecord(
            "incident", 
            caseId, 
            "?$select=incidentid,title&$expand=customerid_account($select=accountid,name),customerid_contact($select=contactid,firstname,lastname,fullname)"
        );
        
        // Check if customer is an Account
        if (caseRecord.customerid_account) {
            const accountId = caseRecord.customerid_account.accountid;
            const accountName = caseRecord.customerid_account.name;
            
            console.log("Case customer is Account:", accountId, accountName);
            
            // Set the duc_subaccount field
            const subAccountField = formContext.getAttribute("duc_subaccount");
            if (subAccountField) {
                subAccountField.setValue([{
                    id: accountId,
                    name: accountName,
                    entityType: "account"
                }]);
                
                // Fire onChange event
                subAccountField.fireOnChange();
                
                console.log("duc_subaccount set to:", accountName);
            }
        }
        // Check if customer is a Contact
        else if (caseRecord.customerid_contact) {
            const contactId = caseRecord.customerid_contact.contactid;
            const contactFullName = caseRecord.customerid_contact.fullname || 
                                   `${caseRecord.customerid_contact.firstname || ''} ${caseRecord.customerid_contact.lastname || ''}`.trim();
            
            console.log("Case customer is Contact:", contactId, contactFullName);
            
            // Create a dummy account with the contact's name
            const dummyAccountId = await createDummyAccountForContact(contactId, contactFullName);
            
            if (dummyAccountId) {
                // Set the duc_subaccount field with the newly created account
                const subAccountField = formContext.getAttribute("duc_subaccount");
                if (subAccountField) {
                    subAccountField.setValue([{
                        id: dummyAccountId,
                        name: contactFullName,
                        entityType: "account"
                    }]);
                    
                    // Fire onChange event
                    subAccountField.fireOnChange();
                    
                    console.log("duc_subaccount set to dummy account:", contactFullName);
                }
            }
        }
        else {
            console.log("Case has no customer (neither Account nor Contact)");
        }
        
    } catch (error) {
        console.error("Error in onLoadHandler:", error);
        Xrm.Navigation.openErrorDialog({ message: "Error processing service request: " + error.message });
    }
}

/**
 * Creates a dummy account for a contact
 * @param {string} contactId - The contact ID
 * @param {string} contactName - The contact's full name
 * @returns {Promise<string>} - The ID of the created account
 */
async function createDummyAccountForContact(contactId, contactName) {
    try {
        // Check if a dummy account already exists for this contact
        const existingAccounts = await Xrm.WebApi.retrieveMultipleRecords(
            "account",
            `?$select=accountid,name&$filter=primarycontactid/contactid eq ${contactId} and name eq '${contactName.replace(/'/g, "''")}'&$top=1`
        );
        
        if (existingAccounts && existingAccounts.entities.length > 0) {
            console.log("Dummy account already exists:", existingAccounts.entities[0].accountid);
            return existingAccounts.entities[0].accountid;
        }
        
        // Create new dummy account
        const accountData = {
            name: contactName,
            "primarycontactid@odata.bind": `/contacts(${contactId})`
            // Add any other fields you need for the dummy account
            // For example: duc_isdummyaccount: true (if you have such a field)
        };
        
        const createdAccount = await Xrm.WebApi.createRecord("account", accountData);
        console.log("Dummy account created:", createdAccount.id);
        
        return createdAccount.id;
        
    } catch (error) {
        console.error("Error creating dummy account:", error);
        throw error;
    }
}