// Mobile function file for Pexip add-in
Office.initialize = function() {};
 
function insertPexipMeeting(event) {
    // Hardcoded VMR details
    var vmrUsername = 'peter.jagtboe';
    var vmrDomain = 'pexip.vc';
    
    var webLink = 'https://' + vmrDomain + '/' + vmrUsername;
    var vcEndpoint = vmrUsername + '@' + vmrDomain;
    var appLink = 'pexip://' + vmrUsername + '@' + vmrDomain;
    
    // Simple text format
    var pexipText = '\n\nPEXIP MEETING DETAILS\n\n' +
        'Join from web browser:\n' + webLink + '\n\n' +
        'Join from VC endpoint:\n' + vcEndpoint + '\n\n' +
        'Join from Pexip App:\n' + appLink + '\n';
    
    // Try to insert as plain text first
    try {
        Office.context.mailbox.item.body.setAsync(
            pexipText,
            { coercionType: Office.CoercionType.Text, asyncContext: { isAppend: true } },
            function(result) {
                if (event && event.completed) {
                    event.completed();
                }
            }
        );
    } catch (e) {
        if (event && event.completed) {
            event.completed();
        }
    }
}
