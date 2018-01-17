//Setting up some global variables for items that would be used in multiple script files.

function getGlobalFormObj() {
    var obj = {
    formId: 'FORMID'
    }
    return obj
}

function testing() {
    var global = getGlobalFormObj();
    Logger.log(global.formId);
}
