async function OnloadCaseIntegration(executionContext) {
    try {
        const formContext = executionContext.getFormContext();
        
        // --- NEW EDIT: Hide new_campaign if service request has a value ---
        const serviceRequestField = formContext.getAttribute("msdyn_servicerequest");
        const campaignControl = formContext.getControl("new_campaign");
        
        if (serviceRequestField && serviceRequestField.getValue()) {
            if (campaignControl) campaignControl.setVisible(false);
        } else {
            if (campaignControl) campaignControl.setVisible(true);
        }
        // -----------------------------------------------------------------

        if (!serviceRequestField) {
            console.log("Service request field not found");
            return;
        }
        
        const serviceRequestValue = serviceRequestField.getValue();
        if (!serviceRequestValue || serviceRequestValue.length === 0) {
            console.log("No service request selected");
            return;
        }
        
        const caseId = serviceRequestValue[0].id.replace(/[{}]/g, '');
        const caseRecord = await Xrm.WebApi.retrieveRecord(
            "incident", 
            caseId, 
            "?$select=incidentid,title,_customerid_value,mme_building,mme_street,mme_zone,mme_municipalitynameen,mme_municipalitynamear,mme_x,mme_y"
        );
        
        if (!caseRecord._customerid_value) {
            console.log("Case has no customer");
            return;
        }
        
        const customerId = caseRecord._customerid_value;
        const customerName = caseRecord["_customerid_value@OData.Community.Display.V1.FormattedValue"];
        const customerType = caseRecord["_customerid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
        
        const subAccountField = formContext.getAttribute("duc_subaccount");
        let finalAccountId = null;
        let finalAccountName = customerName;
        
        if (customerType === "account") {
            finalAccountId = customerId;
            subAccountField.setValue([{ id: customerId, name: customerName, entityType: "account" }]);
            subAccountField.fireOnChange();
        } 
        else if (customerType === "contact") {
            const accountId = await findOrCreateDummyAccount(customerId, customerName);
            if (accountId) {
                finalAccountId = accountId;
                subAccountField.setValue([{ id: accountId, name: customerName, entityType: "account" }]);
                subAccountField.fireOnChange();
            }
        }

        // --- UPDATED: Create address and set the duc_address lookup ---
        if (finalAccountId) {
            const addressInfo = await createAddressInformation(finalAccountId, finalAccountName, caseRecord);
            if (addressInfo) {
                const addressLookupField = formContext.getAttribute("duc_address");
                if (addressLookupField) {
                    addressLookupField.setValue([{
                        id: addressInfo.id,
                        name: addressInfo.name,
                        entityType: "duc_addressinformation"
                    }]);
                    addressLookupField.fireOnChange();
                }
            }
        }
        
    } catch (error) {
        console.error("Error in onLoadHandler:", error);
        Xrm.Navigation.openErrorDialog({ message: "Error processing service request: " + error.message });
    }
}


async function createAddressInformation(accountId, accountName, caseRecord) {
    try {
        const today = new Date();
        const formattedDate = today.toISOString().split('T')[0];
        const addressName = `${accountName} ${formattedDate}`;
        
        const addressData = {
            duc_name: addressName,
            duc_buildingnumber: caseRecord.mme_building || null,
            duc_streetnumber: caseRecord.mme_street || null,
            duc_zonenamear: caseRecord.mme_zone || null,
            duc_address: `${caseRecord.mme_municipalitynameen || ''} ${caseRecord.mme_municipalitynamear || ''}`.trim() || null,
            duc_latitude: caseRecord.mme_y || null,
            duc_longitude: caseRecord.mme_x || null,
            'duc_Account@odata.bind': `/accounts(${accountId.replace(/[{}]/g, '')})`
        };
        
        Object.keys(addressData).forEach(key => {
            if (addressData[key] === null || addressData[key] === '') delete addressData[key];
        });
        
        const createdAddress = await Xrm.WebApi.createRecord('duc_addressinformation', addressData);
        console.log('Address information created:', createdAddress.id);
        
        // Return both ID and Name so it can be used for the Lookup field
        return {
            id: createdAddress.id,
            name: addressName
        };
        
    } catch (error) {
        console.error('Error creating address information:', error);
        return null; 
    }
}

async function findOrCreateDummyAccount(contactId, contactName) {
    try {
        // Remove curly braces if present
        const cleanContactId = contactId.replace(/[{}]/g, '');
        
        // Search for existing account with BOTH the identifier AND primary contact
        // This ensures we find the right account and avoid duplicates
        const existingAccounts = await Xrm.WebApi.retrieveMultipleRecords(
            "account",
            `?$select=accountid,name&$filter=_primarycontactid_value eq ${cleanContactId}&$top=1`
        );
        
        // If account exists, return it
        if (existingAccounts && existingAccounts.entities.length > 0) {
            console.log("Account found with identifier and contact:", existingAccounts.entities[0].accountid);
            return existingAccounts.entities[0].accountid;
        }
        
        // Account doesn't exist, create new one
        const accountData = {
            name: contactName,
            duc_accountidentifier: contactName,
            "PrimaryContactId@odata.bind": `/contacts(${cleanContactId})`
        };
        
        const createdAccount = await Xrm.WebApi.createRecord("account", accountData);
        console.log("New account created with identifier:", createdAccount.id);
        
        return createdAccount.id;
        
    } catch (error) {
        console.error("Error in findOrCreateDummyAccount:", error);
        throw error;
    }
}

