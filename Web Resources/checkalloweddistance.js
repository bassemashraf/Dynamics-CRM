// navigate to the relevant services task 
function oFST() { try { var f = Xrm.Page, w = f.data.entity.getId(); if (!w) return Xrm.Navigation.openAlertDialog({ text: "Work Order ID not found." }); w = w.replace(/[{}]/g, ""); Xrm.WebApi.retrieveMultipleRecords("msdyn_workorderservicetask", "?$select=msdyn_workorderservicetaskid&$filter=_msdyn_workorder_value eq " + w + "&$orderby=createdon asc&$top=1").then(r => { if (!r.entities.length) return Xrm.Navigation.openAlertDialog({ text: "No related service tasks found." }); Xrm.Navigation.openForm({ entityName: "msdyn_workorderservicetask", entityId: r.entities[0].msdyn_workorderservicetaskid }); }); } catch (e) { } }

async function uLA() {
    try {
        const woId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");

        var woRecord = await Xrm.WebApi.retrieveRecord(
            "msdyn_workorder",
            woId,
            "?$select=_duc_processextension_value"
        );

        if (!woRecord._duc_processextension_value) return;

        var peId = woRecord._duc_processextension_value;

        var peRecord = await Xrm.WebApi.retrieveRecord(
            "duc_processextension",
            peId,
            "?$select=_duc_processdefinition_value,_duc_currentstage_value"
        );

        if (!peRecord._duc_processdefinition_value) return;

        var processDefinitionId = peRecord._duc_processdefinition_value;
        var currentStageId = peRecord._duc_currentstage_value;

        var stageActions = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_stageaction",
            `?$select=duc_stageactionid,_duc_defaultstatus_value
              &$filter=duc_canbetriggeredbytarget eq true
              and _duc_process_value eq ${processDefinitionId}
              and _duc_relatedstage_value eq ${currentStageId}
              &$top=10`
        );

        var actionIdToSet = null;

        for (let item of stageActions.entities) {
            let statusId = item._duc_defaultstatus_value;

            if (statusId) {
                let status = await Xrm.WebApi.retrieveRecord(
                    "duc_processstatus",
                    statusId,
                    "?$select=duc_value"
                );

                if (status.duc_value == 690970002) {
                    actionIdToSet = item.duc_stageactionid;
                    break;
                }
            }
        }

        if (!actionIdToSet) return;

        await Xrm.WebApi.updateRecord("duc_processextension", peId, {
            "duc_LastActionTaken_duc_ProcessExtension@odata.bind":
                `/duc_stageactions(${actionIdToSet})`
        });

    } catch (e) {
        console.warn("uLA error:", e);
    }
}

async function sBIP() {
    try {
        const woId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");
        var bookingResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresourcebooking",
            `?$select=bookableresourcebookingid
              &$filter=_msdyn_workorder_value eq ${woId}
              &$orderby=createdon asc
              &$top=1`
        );

        if (!bookingResult.entities.length) return false;

        var bookingId = bookingResult.entities[0].bookableresourcebookingid;

        var statusResult = await Xrm.WebApi.retrieveMultipleRecords(
            "bookingstatus",
            `?$select=bookingstatusid&$filter=name eq 'In Progress'`
        );

        if (!statusResult.entities.length) {
            Xrm.Navigation.openAlertDialog({ text: "Booking Status 'In Progress' not found." });
            return false;
        }

        var statusId = statusResult.entities[0].bookingstatusid;

        await Xrm.WebApi.updateRecord("bookableresourcebooking", bookingId, {
            "BookingStatus@odata.bind": `/bookingstatuses(${statusId})`
        });

        return true;

    } catch (e) {
        Xrm.Navigation.openAlertDialog({ text: e.message });
        return false;
    }
}

function IsFromMobile() {
    if (Xrm.Page.context.client.getClient() == "Mobile"
        && (Xrm.Page.context.client.getFormFactor() == 2 || Xrm.Page.context.client.getFormFactor() == 3)
    ) {
        return true;
    }
    return false;
}

async function createDailyInspection() {
    try {
        Xrm.Utility.showProgressIndicator("Creating inspection record...");

        const workOrderId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");

        if (!workOrderId) {
            Xrm.Navigation.openAlertDialog({
                text: "Unable to retrieve Work Order ID. Please save the form first."
            });
            Xrm.Utility.closeProgressIndicator();
            return;
        }

        const createdBy = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

        const resourceRes = await Xrm.WebApi.retrieveMultipleRecords(
            "bookableresource",
            `?$select=bookableresourceid&$filter=_userid_value eq ${createdBy} and resourcetype eq 3`
        );

        if (!resourceRes.entities || resourceRes.entities.length === 0) {
            Xrm.Navigation.openAlertDialog({
                text: "No bookable resource found for the current user."
            });
            Xrm.Utility.closeProgressIndicator();
            return;
        }

        const resourceId = resourceRes.entities[0].bookableresourceid;

        const currentDateTime = new Date();

        // Get today's date without time
        const todayStart = new Date(currentDateTime.getFullYear(), currentDateTime.getMonth(), currentDateTime.getDate());

        // Check if attendance record exists for today for this user
        const attendanceRes = await Xrm.WebApi.retrieveMultipleRecords(
            "duc_attendance",
            `?$select=duc_attendanceid&$filter=_duc_user_value eq ${createdBy} and Microsoft.Dynamics.CRM.On(PropertyName='createdon',PropertyValue='${todayStart.toISOString()}')`
        );

        let attendanceId;

        if (attendanceRes.entities && attendanceRes.entities.length > 0) {
            // Use existing attendance record
            attendanceId = attendanceRes.entities[0].duc_attendanceid;
            console.log("Using existing attendance record: " + attendanceId);
        } else {
            // Create new attendance record
            const attendanceData = {
                "duc_User@odata.bind": `/systemusers(${createdBy})`
            };

            const attendanceResult = await Xrm.WebApi.createRecord("duc_attendance", attendanceData);
            attendanceId = attendanceResult.id;
            console.log("Created new attendance record: " + attendanceId);
        }

        // Create daily inspection record with attendance lookup
        const recordData = {
            "duc_WorkOrder@odata.bind": `/msdyn_workorders(${workOrderId})`,
            "duc_BookableResource@odata.bind": `/bookableresources(${resourceId})`,
            "duc_Attendance@odata.bind": `/duc_attendances(${attendanceId})`,
            "duc_startinspectiontime": currentDateTime.toISOString()
        };

        const result = await Xrm.WebApi.createRecord("duc_dailyinspectorinspections", recordData);

        // Safely refresh the form if available
        if (Xrm.Page && Xrm.Page.data && Xrm.Page.data.refresh) {
            Xrm.Page.data.refresh(false);
        }

        console.log("Daily inspection record created with ID: " + result.id);

        Xrm.Utility.closeProgressIndicator();

    } catch (error) {
        Xrm.Utility.closeProgressIndicator();

        Xrm.Navigation.openAlertDialog({
            text: "Error creating daily inspection record: " + error.message
        });

        console.error("Error creating daily inspection:", error);
    }
}

async function hasWorkOrderServiceTask(workOrderId) {
    if (!workOrderId) {
        return false;
    }

    const cleanWorkOrderId = workOrderId.replace(/[{}]/g, "");

    const result = await Xrm.WebApi.retrieveMultipleRecords(
        "msdyn_workorderservicetask",
        `?$select=msdyn_workorderservicetaskid&$filter=_msdyn_workorder_value eq ${cleanWorkOrderId}&$top=1`
    );

    return result.entities && result.entities.length > 0;
}

// function WO_CheckAllowedDistance() {
//     var lang = Xrm.Utility.getGlobalContext().userSettings.languageId;
//     var MSG = (lang === 1025) ? {
//         cannotGetLoc: "تعذر استرداد الموقع الحالي، يرجى مراجعة إعدادات الموقع في الهاتف.",
//         checking: "جاري التحقق من الموقع...",
//         outOfRange: "الموقع الحالي يبعد بمسافة {0} متر عن موقع التفتيش المسجل وأقصى مسافة مسموحة هي {1}.",
//         outOfRangeTitle: "خارج المسافة المسموحة",
//         updateError: "فشل تحديث الموقع على أمر العمل."
//     } : {
//         cannotGetLoc: "Can't get current location. Please check your mobile location settings.",
//         checking: "Checking location...",
//         outOfRange: "The current location is {0} meters away from the registered location and max allowed distance is {1}.",
//         outOfRangeTitle: "Out of Range",
//         updateError: "Failed to update location on work order."
//     };

//     Xrm.Utility.showProgressIndicator(MSG.checking);

//     var workorderId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");

//     return Xrm.WebApi.retrieveRecord("msdyn_workorder", workorderId,
//         "?$select=duc_details_islandmark,duc_details_alloweddistance,duc_channel,_duc_department_value"
//     ).then(function (workOrder) {

//         var allowedDist = null;
//         var isLandmark = workOrder.duc_details_islandmark;
//         var isBSS = workOrder.duc_channel === 100000002;

//         if (isLandmark && workOrder.duc_details_alloweddistance) {
//             allowedDist = workOrder.duc_details_alloweddistance;
//             return proceedWithDistanceCheck(allowedDist);
//         }
//         else if (workOrder._duc_department_value) {
//             var ouId = workOrder._duc_department_value.replace(/[{}]/g, "");
//             var fieldName = isBSS ? "duc_bssdistanceinmeter" : "duc_distanceinmeter";

//             return Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", ouId,
//                 "?$select=" + fieldName
//             ).then(function (ou) {
//                 if (ou[fieldName] != null) {
//                     allowedDist = ou[fieldName] - 100000000;
//                 } else {
//                     allowedDist = 0;
//                 }
//                 return proceedWithDistanceCheck(allowedDist);
//             }).catch(function () {
//                 allowedDist = 0;
//                 return proceedWithDistanceCheck(allowedDist);
//             });
//         }
//         else {
//             allowedDist = 0;
//             return proceedWithDistanceCheck(allowedDist);
//         }

//     }).catch(function (error) {
//         Xrm.Utility.closeProgressIndicator();
//         Xrm.Navigation.openAlertDialog({
//             text: (error && error.message) ? error.message : "Failed to read work order."
//         });
//         return Promise.resolve(false); // Return false on error
//     });

//     function proceedWithDistanceCheck(allowedDistance) {
//         if (!allowedDistance || allowedDistance <= 0) {
//             return getCurrentLocationAndUpdate(null, null, null);
//         }

//         var addrAttr = Xrm.Page.getAttribute("duc_address");
//         var addrVal = addrAttr ? addrAttr.getValue() : null;
//         if (!addrVal || !addrVal[0] || !addrVal[0].id) {
//             return getCurrentLocationAndUpdate(null, null, allowedDistance);
//         }

//         var addressId = addrVal[0].id.replace(/[{}]/g, "");
//         var addressEntity = "duc_addressinformation";

//         return Xrm.WebApi.retrieveRecord(addressEntity, addressId, "?$select=duc_latitude,duc_longitude").then(
//             function (result) {
//                 var destLat = result.duc_latitude;
//                 var destLng = result.duc_longitude;

//                 if (destLat == null || destLng == null) {
//                     return getCurrentLocationAndUpdate(null, null, allowedDistance);
//                 }

//                 return getCurrentLocationAndUpdate(destLat, destLng, allowedDistance);
//             },
//             function (error) {
//                 Xrm.Utility.closeProgressIndicator();
//                 Xrm.Navigation.openAlertDialog({
//                     text: (error && error.message) ? error.message : "Failed to read address record."
//                 });
//                 return Promise.resolve(false); // Return false on error
//             }
//         );
//     }

//     function getCurrentLocationAndUpdate(destLat, destLng, allowedDistance) {
//         return Xrm.Device.getCurrentPosition().then(function (location) {
//             var originLat = location.coords.latitude;
//             var originLng = location.coords.longitude;

//             if (destLat != null && destLng != null && allowedDistance != null && allowedDistance > 0) {
//                 var distance = GetDistance(originLat, originLng, destLat, destLng);

//                 if (distance > allowedDistance) {
//                     Xrm.Utility.closeProgressIndicator();

//                     var errorMsg = MSG.outOfRange.replace("{0}", Math.round(distance)).replace("{1}", allowedDistance);

//                     Xrm.Navigation.openAlertDialog({
//                         title: MSG.outOfRangeTitle,
//                         text: errorMsg
//                     });

//                     return Promise.resolve(false); // Return false - out of range
//                 }
//             }

//             return updateWorkOrderLocation(originLat, originLng);

//         }, function () {
//             Xrm.Utility.closeProgressIndicator();
//             Xrm.Navigation.openAlertDialog({ text: MSG.cannotGetLoc });
//             return Promise.resolve(false); // Return false on location error
//         });
//     }

//     function updateWorkOrderLocation(originLat, originLng) {
//         var updateData = {
//             "msdyn_latitude": originLat,
//             "msdyn_longitude": originLng
//         };

//         return Xrm.WebApi.updateRecord("msdyn_workorder", workorderId, updateData).then(
//             function success(result) {
//                 Xrm.Utility.closeProgressIndicator();

//                 // Safely refresh the form if available
//                 if (Xrm.Page && Xrm.Page.data && Xrm.Page.data.refresh) {
//                     Xrm.Page.data.refresh(false);
//                 }

//                 return Promise.resolve(true); // Return true - success
//             },
//             function (error) {
//                 Xrm.Utility.closeProgressIndicator();
//                 Xrm.Navigation.openAlertDialog({
//                     text: MSG.updateError + "\n" + ((error && error.message) ? error.message : "")
//                 });
//                 return Promise.resolve(false); // Return false on update error
//             }
//         );
//     }
// }

// function GetDistance(lat1, lon1, lat2, lon2) {
//     var R = 6371000;
//     var dLat = (lat2 - lat1) * Math.PI / 180;
//     var dLon = (lon2 - lon1) * Math.PI / 180;
//     var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
//         Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
//         Math.sin(dLon / 2) * Math.sin(dLon / 2);
//     var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
//     var distance = R * c;
//     return distance;
// }

// ============================================
// MAIN EXECUTION - ONLY CHECK DISTANCE ON MOBILE
// ============================================



function WO_CheckAllowedDistance() {
    var lang = Xrm.Utility.getGlobalContext().userSettings.languageId;
    var MSG = (lang === 1025) ? {
        cannotGetLoc: "تعذر استرداد الموقع الحالي، يرجى مراجعة إعدادات الموقع في الهاتف.",
        checking: "جاري التحقق من الموقع...",
        outOfRange: "الموقع الحالي يبعد بمسافة {0} متر عن موقع التفتيش المسجل وأقصى مسافة مسموحة هي {1}.",
        outOfRangeTitle: "خارج المسافة المسموحة",
        updateError: "فشل تحديث الموقع على أمر العمل.",
        noSubAccount: "لا يوجد حساب فرعي محدد في أمر العمل.",
        noAddresses: "لم يتم العثور على عناوين مرتبطة بالحساب الفرعي."
    } : {
        cannotGetLoc: "Can't get current location. Please check your mobile location settings.",
        checking: "Checking location...",
        outOfRange: "The current location is {0} meters away from the registered location and max allowed distance is {1}.",
        outOfRangeTitle: "Out of Range",
        updateError: "Failed to update location on work order.",
        noSubAccount: "No sub-account is selected on the Work Order.",
        noAddresses: "No addresses found linked to the sub-account."
    };

    Xrm.Utility.showProgressIndicator(MSG.checking);

    var workorderId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");

    return Xrm.WebApi.retrieveRecord("msdyn_workorder", workorderId,
        "?$select=duc_details_islandmark,duc_details_alloweddistance,duc_channel,_duc_department_value"
    ).then(function (workOrder) {

        var allowedDist = null;
        var isLandmark = workOrder.duc_details_islandmark;
        var isBSS = workOrder.duc_channel === 100000002;

        if (isLandmark && workOrder.duc_details_alloweddistance) {
            allowedDist = workOrder.duc_details_alloweddistance;
            return proceedWithDistanceCheck(allowedDist);
        }
        else if (workOrder._duc_department_value) {
            var ouId = workOrder._duc_department_value.replace(/[{}]/g, "");
            var fieldName = isBSS ? "duc_bssdistanceinmeter" : "duc_distanceinmeter";

            return Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", ouId,
                "?$select=" + fieldName
            ).then(function (ou) {
                if (ou[fieldName] != null) {
                    allowedDist = ou[fieldName] - 100000000;
                } else {
                    allowedDist = 0;
                }
                return proceedWithDistanceCheck(allowedDist);
            }).catch(function () {
                allowedDist = 0;
                return proceedWithDistanceCheck(allowedDist);
            });
        }
        else {
            allowedDist = 0;
            return proceedWithDistanceCheck(allowedDist);
        }

    }).catch(function (error) {
        Xrm.Utility.closeProgressIndicator();
        Xrm.Navigation.openAlertDialog({
            text: (error && error.message) ? error.message : "Failed to read work order."
        });
        return Promise.resolve(false);
    });

    function proceedWithDistanceCheck(allowedDistance) {
        if (!allowedDistance || allowedDistance <= 0) {
            return getCurrentLocationAndUpdate(null, null, null);
        }

        // First, try to get address from duc_address lookup
        var addrAttr = Xrm.Page.getAttribute("duc_address");
        var addrVal = addrAttr ? addrAttr.getValue() : null;
        
        if (addrVal && addrVal[0] && addrVal[0].id) {
            var addressId = addrVal[0].id.replace(/[{}]/g, "");
            var addressEntity = "duc_addressinformation";

            return Xrm.WebApi.retrieveRecord(addressEntity, addressId, "?$select=duc_latitude,duc_longitude").then(
                function (result) {
                    var destLat = result.duc_latitude;
                    var destLng = result.duc_longitude;

                    if (destLat != null && destLng != null) {
                        return getCurrentLocationAndUpdate(destLat, destLng, allowedDistance);
                    } else {
                        // If address doesn't have coordinates, fall back to account addresses
                        return checkNearestAccountAddress(allowedDistance);
                    }
                },
                function (error) {
                    // If error retrieving address, fall back to account addresses
                    return checkNearestAccountAddress(allowedDistance);
                }
            );
        } else {
            // No duc_address lookup, check account addresses
            return checkNearestAccountAddress(allowedDistance);
        }
    }

    function checkNearestAccountAddress(allowedDistance) {
        // Get the sub-account from the work order
        var subAccountAttr = Xrm.Page.getAttribute("duc_subaccount");
        var subAccountVal = subAccountAttr ? subAccountAttr.getValue() : null;
        
        if (!subAccountVal || !subAccountVal[0] || !subAccountVal[0].id) {
            // No sub-account, proceed without distance check
            return getCurrentLocationAndUpdate(null, null, null);
        }

        var accountId = subAccountVal[0].id.replace(/[{}]/g, "");

        // Build OData query to fetch addresses linked to the account
        var options = "?$select=duc_addressinformationid,duc_latitude,duc_longitude";
        options += "&$filter=_duc_account_value eq " + accountId;
        options += " and duc_latitude ne null and duc_longitude ne null";

        return Xrm.WebApi.retrieveMultipleRecords("duc_addressinformation", options).then(
            function (result) {
                if (!result.entities || result.entities.length === 0) {
                    // No addresses found, proceed without distance check
                    return getCurrentLocationAndUpdate(null, null, null);
                }

                // Filter addresses with valid coordinates
                var addresses = result.entities.filter(function(addr) {
                    return addr.duc_latitude != null && addr.duc_longitude != null;
                });

                if (addresses.length === 0) {
                    // No valid addresses, proceed without distance check
                    return getCurrentLocationAndUpdate(null, null, null);
                }

                // Get current location first to find nearest address
                return Xrm.Device.getCurrentPosition().then(function (location) {
                    var originLat = location.coords.latitude;
                    var originLng = location.coords.longitude;

                    // Calculate distance to each address and find the nearest
                    var nearestAddress = null;
                    var minDistance = Infinity;

                    addresses.forEach(function(addr) {
                        var distance = GetDistance(originLat, originLng, addr.duc_latitude, addr.duc_longitude);
                        if (distance < minDistance) {
                            minDistance = distance;
                            nearestAddress = {
                                lat: addr.duc_latitude,
                                lng: addr.duc_longitude,
                                distance: distance
                            };
                        }
                    });

                    // Check if nearest address is within allowed distance
                    if (nearestAddress && allowedDistance > 0 && nearestAddress.distance > allowedDistance) {
                        Xrm.Utility.closeProgressIndicator();

                        var errorMsg = MSG.outOfRange
                            .replace("{0}", Math.round(nearestAddress.distance))
                            .replace("{1}", allowedDistance);

                        Xrm.Navigation.openAlertDialog({
                            title: MSG.outOfRangeTitle,
                            text: errorMsg
                        });

                        return Promise.resolve(false);
                    }

                    // Within range or no distance check needed, update work order
                    return updateWorkOrderLocation(originLat, originLng);

                }, function () {
                    Xrm.Utility.closeProgressIndicator();
                    Xrm.Navigation.openAlertDialog({ text: MSG.cannotGetLoc });
                    return Promise.resolve(false);
                });
            },
            function (error) {
                // Error retrieving addresses, proceed without distance check
                return getCurrentLocationAndUpdate(null, null, null);
            }
        );
    }

    function getCurrentLocationAndUpdate(destLat, destLng, allowedDistance) {
        return Xrm.Device.getCurrentPosition().then(function (location) {
            var originLat = location.coords.latitude;
            var originLng = location.coords.longitude;

            if (destLat != null && destLng != null && allowedDistance != null && allowedDistance > 0) {
                var distance = GetDistance(originLat, originLng, destLat, destLng);

                if (distance > allowedDistance) {
                    Xrm.Utility.closeProgressIndicator();

                    var errorMsg = MSG.outOfRange.replace("{0}", Math.round(distance)).replace("{1}", allowedDistance);

                    Xrm.Navigation.openAlertDialog({
                        title: MSG.outOfRangeTitle,
                        text: errorMsg
                    });

                    return Promise.resolve(false);
                }
            }

            return updateWorkOrderLocation(originLat, originLng);

        }, function () {
            Xrm.Utility.closeProgressIndicator();
            Xrm.Navigation.openAlertDialog({ text: MSG.cannotGetLoc });
            return Promise.resolve(false);
        });
    }

    function updateWorkOrderLocation(originLat, originLng) {
        var updateData = {
            "msdyn_latitude": originLat,
            "msdyn_longitude": originLng
        };

        return Xrm.WebApi.updateRecord("msdyn_workorder", workorderId, updateData).then(
            function success(result) {
                Xrm.Utility.closeProgressIndicator();

                if (Xrm.Page && Xrm.Page.data && Xrm.Page.data.refresh) {
                    Xrm.Page.data.refresh(false);
                }

                return Promise.resolve(true);
            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
                Xrm.Navigation.openAlertDialog({
                    text: MSG.updateError + "\n" + ((error && error.message) ? error.message : "")
                });
                return Promise.resolve(false);
            }
        );
    }
}

function GetDistance(lat1, lon1, lat2, lon2) {
    var R = 6371000;
    var dLat = (lat2 - lat1) * Math.PI / 180;
    var dLon = (lon2 - lon1) * Math.PI / 180;
    var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
        Math.sin(dLon / 2) * Math.sin(dLon / 2);
    var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    var distance = R * c;
    return distance;
}
(async function () {
    try {
        const workOrderId = Xrm.Page.data.entity.getId();
        const isMobile = IsFromMobile();

        // Step 1: Check if service tasks exist
        const hasTasks = await hasWorkOrderServiceTask(workOrderId);

        if (!hasTasks) {
            Xrm.Navigation.openAlertDialog({
                text: "No Service Tasks found for this Work Order, please try again in few moments"
            });
            return; // Stop execution
        }

        // Step 2: Check distance ONLY if on mobile
        if (isMobile) {
            const distanceCheckPassed = await WO_CheckAllowedDistance();
            
            // Step 3: Only create daily inspection if distance check passed
            if (distanceCheckPassed) {
                await createDailyInspection();
            } else {
                // Distance check failed - stop execution
                return;
            }
        }

        // Step 4: Continue with other functions
        oFST();

        var updateBookableResourcesBooking = await sBIP();

        if (updateBookableResourcesBooking) {
            await uLA();
        }

    } catch (error) {
        // If any error occurs, stop execution
        console.error("Execution stopped due to error:", error);
    }
})();