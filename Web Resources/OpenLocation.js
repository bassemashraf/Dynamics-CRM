function WO_OpenDirectionsToDucAddress() {
    var client = Xrm.Utility.getGlobalContext().client.getClient();
    if (client !== "Mobile") {
        Xrm.Navigation.openAlertDialog({ text: "This action is available only on mobile." });
        return;
    }

    var lang = Xrm.Utility.getGlobalContext().userSettings.languageId;
    var MSG = (lang === 1025) ? { 
        noLookup: "لا يوجد عنوان محدد في أمر العمل (الحقل: duc_address).",
        noCoords: "سجل العنوان لا يحتوي على خطوط الطول/العرض (duc_latitude / duc_longitude).",
        cannotGetLoc: "تعذر استرداد الموقع الحالي، يرجى مراجعة إعدادات الموقع في الهاتف.",
        confirmTitle: "الاتجاهات في خرائط جوجل",
        confirmText: "هل ترغب في فتح خرائط جوجل للحصول على الاتجاهات إلى موقع التفتيش?"
    } : { 
        noLookup: "No address is selected on the Work Order (field: duc_address).",
        noCoords: "The address record does not have latitude/longitude (duc_latitude / duc_longitude).",
        cannotGetLoc: "Can't get current location. Please check your mobile location settings.",
        confirmTitle: "Google Maps Directions",
        confirmText: "Open Google Maps to get directions to the inspection location?"
    };

    var addrAttr = Xrm.Page.getAttribute("duc_address");
    var addrVal = addrAttr ? addrAttr.getValue() : null;
    if (!addrVal || !addrVal[0] || !addrVal[0].id) {
        Xrm.Navigation.openAlertDialog({ text: MSG.noLookup });
        return;
    }

    var addressId = addrVal[0].id.replace(/[{}]/g, "");
    var addressEntity = "duc_addressinformation";

    Xrm.WebApi.retrieveRecord(addressEntity, addressId, "?$select=duc_latitude,duc_longitude").then(
        function (result) {
            var destLat = result.duc_latitude;
            var destLng = result.duc_longitude;

            if (destLat == null || destLng == null) {
                Xrm.Navigation.openAlertDialog({ text: MSG.noCoords });
                return;
            }

            Xrm.Device.getCurrentPosition().then(function (location) {
                var originLat = location.coords.latitude;
                var originLng = location.coords.longitude;

                var url = "https://www.google.com/maps/dir/?api=1"
                    + "&origin=" + originLat + "," + originLng
                    + "&destination=" + destLat + "," + destLng
                    + "&travelmode=driving";

                Xrm.Navigation.openUrl(url);
            }, function () {
                Xrm.Navigation.openAlertDialog({ text: MSG.cannotGetLoc });
            });
        },
        function (error) {
            Xrm.Navigation.openAlertDialog({
                text: (error && error.message) ? error.message : "Failed to read address record."
            });
        }
    );
}
await WO_OpenDirectionsToDucAddress();


function WO_OpenDirectionsToDucAddress() {
    var client = Xrm.Utility.getGlobalContext().client.getClient();
    if (client !== "Mobile") {
        Xrm.Navigation.openAlertDialog({ text: "This action is available only on mobile." });
        return;
    }

    var lang = Xrm.Utility.getGlobalContext().userSettings.languageId;
    var MSG = (lang === 1025) ? { 
        noSubAccount: "لا يوجد حساب فرعي محدد في أمر العمل (الحقل: duc_subaccount).",
        noAddresses: "لم يتم العثور على عناوين مرتبطة بالحساب الفرعي.",
        noCoords: "بعض العناوين لا تحتوي على إحداثيات صالحة.",
        cannotGetLoc: "تعذر استرداد الموقع الحالي، يرجى مراجعة إعدادات الموقع في الهاتف.",
        confirmTitle: "الاتجاهات في خرائط جوجل",
        confirmText: "هل ترغب في فتح خرائط جوجل للحصول على الاتجاهات مع نقاط التوقف المتعددة؟",
        tooManyWaypoints: "تم العثور على أكثر من 25 عنوانًا. سيتم استخدام أقرب 25 موقعًا فقط.",
        loadingMessage: "جاري تحميل العناوين وبناء المسار..."
    } : { 
        noSubAccount: "No sub-account is selected on the Work Order (field: duc_subaccount).",
        noAddresses: "No addresses found linked to the sub-account.",
        noCoords: "Some addresses do not have valid coordinates.",
        cannotGetLoc: "Can't get current location. Please check your mobile location settings.",
        confirmTitle: "Google Maps Directions",
        confirmText: "Open Google Maps to get directions with multiple stops?",
        tooManyWaypoints: "Found more than 25 addresses. Only the nearest 25 locations will be used.",
        loadingMessage: "Loading addresses and building route..."
    };

    var subAccountAttr = Xrm.Page.getAttribute("duc_subaccount");
    var subAccountVal = subAccountAttr ? subAccountAttr.getValue() : null;
    
    if (!subAccountVal || !subAccountVal[0] || !subAccountVal[0].id) {
        Xrm.Navigation.openAlertDialog({ text: MSG.noSubAccount });
        return;
    }

    var accountId = subAccountVal[0].id.replace(/[{}]/g, "");

    function calculateDistance(lat1, lon1, lat2, lon2) {
        var R = 6371; // Earth's radius in km
        var dLat = (lat2 - lat1) * Math.PI / 180;
        var dLon = (lon2 - lon1) * Math.PI / 180;
        var a = Math.sin(dLat/2) * Math.sin(dLat/2) +
                Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
                Math.sin(dLon/2) * Math.sin(dLon/2);
        var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
        return R * c;
    }

    Xrm.Utility.showProgressIndicator(MSG.loadingMessage);

    Xrm.Device.getCurrentPosition().then(function (location) {
        var originLat = location.coords.latitude;
        var originLng = location.coords.longitude;

        var options = "?$select=duc_addressinformationid,duc_latitude,duc_longitude";
        options += "&$filter=_duc_account_value eq " + accountId;
        options += " and duc_latitude ne null and duc_longitude ne null";

        Xrm.WebApi.retrieveMultipleRecords("duc_addressinformation", options).then(
            function (result) {
                if (!result.entities || result.entities.length === 0) {
                    Xrm.Utility.closeProgressIndicator();
                    Xrm.Navigation.openAlertDialog({ text: MSG.noAddresses });
                    return;
                }

                var addresses = result.entities
                    .filter(function(addr) {
                        return addr.duc_latitude != null && addr.duc_longitude != null;
                    })
                    .map(function(addr) {
                        return {
                            lat: addr.duc_latitude,
                            lng: addr.duc_longitude,
                            distance: calculateDistance(originLat, originLng, addr.duc_latitude, addr.duc_longitude)
                        };
                    })
                    .sort(function(a, b) {
                        return a.distance - b.distance; // Sort from nearest to farthest
                    });

                if (addresses.length === 0) {
                    Xrm.Utility.closeProgressIndicator();
                    Xrm.Navigation.openAlertDialog({ text: MSG.noAddresses });
                    return;
                }

                var maxWaypoints = 25;
                if (addresses.length > maxWaypoints) {
                    addresses = addresses.slice(0, maxWaypoints);
                    Xrm.Navigation.openAlertDialog({ text: MSG.tooManyWaypoints });
                }

                var url = "https://www.google.com/maps/dir/?api=1";
                url += "&origin=" + originLat + "," + originLng;
                
                var destination = addresses[addresses.length - 1];
                url += "&destination=" + destination.lat + "," + destination.lng;
                
                if (addresses.length > 1) {
                    var waypoints = addresses.slice(0, -1).map(function(addr) {
                        return addr.lat + "," + addr.lng;
                    }).join("|");
                    url += "&waypoints=" + waypoints;
                }
                
                url += "&travelmode=driving";

                Xrm.Utility.closeProgressIndicator();

                var confirmOptions = {
                    title: MSG.confirmTitle,
                    text: MSG.confirmText + " (" + addresses.length + " " + (lang === 1025 ? "موقع" : "locations") + ")",
                    confirmButtonLabel: lang === 1025 ? "نعم" : "Yes",
                    cancelButtonLabel: lang === 1025 ? "إلغاء" : "Cancel"
                };

                Xrm.Navigation.openConfirmDialog(confirmOptions).then(function(response) {
                    if (response.confirmed) {
                        Xrm.Navigation.openUrl(url);
                    }
                });
            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
                Xrm.Navigation.openAlertDialog({
                    text: (error && error.message) ? error.message : "Failed to retrieve address records."
                });
            }
        );
    }, function () {
        Xrm.Utility.closeProgressIndicator();
        Xrm.Navigation.openAlertDialog({ text: MSG.cannotGetLoc });
    });
}

await WO_OpenDirectionsToDucAddress();