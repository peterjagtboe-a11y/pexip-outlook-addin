// Mobile function file for Pexip add-in
Office.initialize = function() {};

function insertPexipMeeting(event) {
    // Hardcoded VMR details for Peter Jagtboe
    const vmrUsername = 'peter.jagtboe';
    const vmrDomain = 'pexip.vc';
    
    const webLink = 'https://' + vmrDomain + '/' + vmrUsername;
    const vcEndpoint = vmrUsername + '@' + vmrDomain;
    
    // Create HTML formatted content
    const pexipHTML = '<br><br><strong>PEXIP MEETING DETAILS</strong><br><br>' +
        '<strong>Join from web browser:</strong> ' + webLink + '<br><br>' +
        '<strong>Join from VC endpoint:</strong> ' + vcEndpoint + '<br><br>' +
    
    // Insert into appointment body
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const currentBody = result.value || '';
                const newBody = currentBody + pexipHTML;
                
                Office.context.mailbox.item.body.setAsync(
                    newBody,
                    { coercionType: Office.CoercionType.Html },
                    function(setResult) {
                        // Signal completion
                        event.completed();
                    }
                );
            } else {
                event.completed();
            }
        }
    );
}
