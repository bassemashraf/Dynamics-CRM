var formContext = Xrm.Page;

formContext.data.save().then(function () {

    var entityName = formContext.data.entity.getEntityName();
    var entityId = formContext.data.entity.getId();

    Xrm.Navigation.openForm({
        entityName: entityName,
        entityId: entityId
    });


}).catch(function (e) {
    console.log("Save error:", e);
});