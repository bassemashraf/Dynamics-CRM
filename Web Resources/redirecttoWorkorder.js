async function redirectToWorkOrder() {
    var workOrderLookup = await Xrm.Page.getAttribute("msdyn_workorder").getValue();

    if (workOrderLookup && workOrderLookup.length > 0) {
        var workOrderId = workOrderLookup[0].id;
        var workOrderName = workOrderLookup[0].name;

        workOrderId = workOrderId.replace(/[{}]/g, "");

        var pageInput = {
            pageType: "entityrecord",
            entityName: "msdyn_workorder",
            entityId: workOrderId
        };

        var navigationOptions = {
            target: 1,
            position: 1
        };

        Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
            function success() {
                console.log("Navigated to Work Order: " + workOrderName);
            },
            function error(err) {
                Xrm.Navigation.openAlertDialog({
                    text: "Error navigating to work order: " + err.message
                });
            }
        );
    } else {
        Xrm.Navigation.openAlertDialog({
            text: "Work Order Field Is Empty"
        });
    }
}
await redirectToWorkOrder();
