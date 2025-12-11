function hideFormHeader(executionContext) {
    var formContext = executionContext.getFormContext();
    debugger
    // Hide header items
    formContext.ui.headerSection.setBodyVisible(false);
    formContext.ui.headerSection.setCommandBarVisible(false);
    formContext.ui.headerSection.setTabNavigatorVisible(false);

    // Force ribbon update so mobile removes Save / Save & Close buttons
    setTimeout(() => {
        formContext.ui.refreshRibbon();
    }, 200);

    // Optional: Only refresh form on first load (if needed)
    if (executionContext.getDepth() === 1) {
        setTimeout(() => {
            formContext.data.refresh(false);
        }, 400);
    }

    setTimeout(() => {
 document.getElementById('id-87143487-8dd6-4df2-8b12-22719e4f40b0-1-duc_name-field-label').parentElement.style.display = 'none'

    }, 2000);

}