function WO_CheckAllowedDistance() {
    var client = Xrm.Utility.getGlobalContext().client.getClient();
    if (client !== "Mobile") {
        Xrm.Navigation.openAlertDialog({ text: "This action is available only on mobile." });
        return;
    }

    var lang = Xrm.Utility.getGlobalContext().userSettings.languageId;
    var MSG = (lang === 1025) ? { // Arabic
        noLookup: "لا يوجد عنوان محدد في أمر العمل (الحقل: duc_address).",
        noCoords: "سجل العنوان لا يحتوي على خطوط الطول/العرض (duc_latitude / duc_longitude).",
        cannotGetLoc: "تعذر استرداد الموقع الحالي، يرجى مراجعة إعدادات الموقع في الهاتف.",
        checking: "جاري التحقق من الموقع...",
        withinRange: "أنت ضمن المسافة المسموحة! المسافة الحالية: {0} متر (المسموح: {1} متر)",
        outOfRange: "الموقع الحالي يبعد بمسافة {0} متر عن موقع التفتيش المسجل وأقصى مسافة مسموحة هي {1}. هل ترغب في فتح خريطة جوجل لتتمكن من تحديد الموقع؟",
        outOfRangeTitle: "خارج المسافة المسموحة",
        withinRangeTitle: "الموقع صحيح",
        noDistance: "لا توجد مسافة محددة للتحقق"
    } : { // English
        noLookup: "No address is selected on the Work Order (field: duc_address).",
        noCoords: "The address record does not have latitude/longitude (duc_latitude / duc_longitude).",
        cannotGetLoc: "Can't get current location. Please check your mobile location settings.",
        checking: "Checking location...",
        withinRange: "You are within allowed distance! Current distance: {0} meters (allowed: {1} meters)",
        outOfRange: "The current location is {0} meters away from the registered location and max allowed distance is {1}. Would you like to open Google Maps to locate it?",
        outOfRangeTitle: "Out of Range",
        withinRangeTitle: "Location Verified",
        noDistance: "No distance restriction configured"
    };

    // Show progress indicator
    Xrm.Utility.showProgressIndicator(MSG.checking);

    var workorderId = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");

    // Get work order details
    Xrm.WebApi.retrieveRecord("msdyn_workorder", workorderId, 
        "?$select=duc_details_islandmark,duc_details_alloweddistance,duc_channel,_duc_organizationalunitid_value"
    ).then(function(workOrder) {
        
        var allowedDist = null;
        var isLandmark = workOrder.duc_details_islandmark;
        var isBSS = workOrder.duc_channel === 100000002;
        
        // Check if landmark has specific allowed distance
        if (isLandmark && workOrder.duc_details_alloweddistance) {
            allowedDist = workOrder.duc_details_alloweddistance;
            proceedWithDistanceCheck(allowedDist);
        } 
        // Get from organizational unit directly from work order
        else if (workOrder._duc_organizationalunitid_value) {
            var ouId = workOrder._duc_organizationalunitid_value.replace(/[{}]/g, "");
            var fieldName = isBSS ? "duc_bssdistanceinmeter" : "duc_distanceinmeter";
            
            Xrm.WebApi.retrieveRecord("msdyn_organizationalunit", ouId, 
                "?$select=" + fieldName
            ).then(function(ou) {
                if (ou[fieldName] != null) {
                    allowedDist = ou[fieldName] - 100000000;
                } else {
                    allowedDist = getConfig("DistanceInMetre") || 0;
                }
                proceedWithDistanceCheck(allowedDist);
            }).catch(function() {
                allowedDist = getConfig("DistanceInMetre") || 0;
                proceedWithDistanceCheck(allowedDist);
            });
        } 
        else {
            allowedDist = getConfig("DistanceInMetre") || 0;
            proceedWithDistanceCheck(allowedDist);
        }
        
    }).catch(function(error) {
        Xrm.Utility.closeProgressIndicator();
        Xrm.Navigation.openAlertDialog({
            text: (error && error.message) ? error.message : "Failed to read work order."
        });
    });

    function proceedWithDistanceCheck(allowedDistance) {
        if (!allowedDistance || allowedDistance <= 0) {
            Xrm.Utility.closeProgressIndicator();
            Xrm.Navigation.openAlertDialog({ text: MSG.noDistance });
            return;
        }

        // Get duc_address lookup value
        var addrAttr = Xrm.Page.getAttribute("duc_address");
        var addrVal = addrAttr ? addrAttr.getValue() : null;
        if (!addrVal || !addrVal[0] || !addrVal[0].id) {
            Xrm.Utility.closeProgressIndicator();
            Xrm.Navigation.openAlertDialog({ text: MSG.noLookup });
            return;
        }

        var addressId = addrVal[0].id.replace(/[{}]/g, "");
        var addressEntity = "duc_addressinformation";

        // Get address coordinates
        Xrm.WebApi.retrieveRecord(addressEntity, addressId, "?$select=duc_latitude,duc_longitude").then(
            function (result) {
                var destLat = result.duc_latitude;
                var destLng = result.duc_longitude;

                if (destLat == null || destLng == null) {
                    Xrm.Utility.closeProgressIndicator();
                    Xrm.Navigation.openAlertDialog({ text: MSG.noCoords });
                    return;
                }

                // Get current location
                Xrm.Device.getCurrentPosition().then(function (location) {
                    var originLat = location.coords.latitude;
                    var originLng = location.coords.longitude;

                    // Calculate distance
                    var distance = GetDistance(originLat, originLng, destLat, destLng);
                    
                    Xrm.Utility.closeProgressIndicator();

                    if (distance <= allowedDistance) {
                        // Within allowed distance
                        logLocation(originLat, originLng, destLat, destLng, distance, true, allowedDistance, workorderId);
                        
                        var successMsg = MSG.withinRange.replace("{0}", Math.round(distance)).replace("{1}", allowedDistance);
                        Xrm.Navigation.openAlertDialog({ 
                            title: MSG.withinRangeTitle,
                            text: successMsg 
                        });
                    } else {
                        // Outside allowed distance
                        logLocation(originLat, originLng, destLat, destLng, distance, false, allowedDistance, workorderId);
                        
                        var errorMsg = MSG.outOfRange.replace("{0}", Math.round(distance)).replace("{1}", allowedDistance);
                        var confirmStrings = { 
                            title: MSG.outOfRangeTitle, 
                            text: errorMsg 
                        };
                        
                        Xrm.Navigation.openConfirmDialog(confirmStrings).then(function (res) {
                            if (res.confirmed) {
                                var url = "https://www.google.com/maps/dir/?api=1"
                                    + "&origin=" + originLat + "," + originLng
                                    + "&destination=" + destLat + "," + destLng
                                    + "&travelmode=driving";
                                Xrm.Navigation.openUrl(url);
                            }
                        });
                    }
                }, function () {
                    Xrm.Utility.closeProgressIndicator();
                    Xrm.Navigation.openAlertDialog({ text: MSG.cannotGetLoc });
                });
            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
                Xrm.Navigation.openAlertDialog({
                    text: (error && error.message) ? error.message : "Failed to read address record."
                });
            }
        );
    }
}

// GetDistance function (Haversine formula)
function GetDistance(lat1, lon1, lat2, lon2) {
    var R = 6371000; // Radius of Earth in meters
    var dLat = (lat2 - lat1) * Math.PI / 180;
    var dLon = (lon2 - lon1) * Math.PI / 180;
    var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
    var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    var distance = R * c;
    return distance;
}