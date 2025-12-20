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