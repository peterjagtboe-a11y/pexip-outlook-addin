// Mobile function file for Pexip add-in
// This executes directly when the button is clicked on mobile

Office.onReady();

function insertPexipMeeting(event) {
    // Hardcoded VMR details for Peter Jagtboe
    const vmrUsername = 'peter.jagtboe';
    const vmrDomain = 'pexip.vc';
    
    const webLink = 'https://' + vmrDomain + '/' + vmrUsername;
    const vcEndpoint = vmrUsername + '@' + vmrDomain;
    const appLink = 'pexip://' + vmrUsername + '@' + vmrDomain;
    
    // Create HTML formatted content
    const pexipHTML = `<br><br><strong>PEXIP MEETING DETAILS</strong><br><br>` +
        `<strong>Join from web browser:</strong> ${webLink}<br><br>` +
        `<strong>Join from VC endpoint:</strong> ${vcEndpoint}<br><br>` +
        `<strong>Join from Pexip App:</strong> ${appLink}<br>`;
    
    // Insert into appointment body
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                let currentBody = result.value || '';
                let newBody = currentBody + pexipHTML;
                
                Office.context.mailbox.item.body.setAsync(
                    newBody,
                    { coercionType: Office.CoercionType.Html },
                    function(setResult) {
                        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                            // Success - must call event.completed() for mobile
                            event.completed();
                        } else {
                            // Error
                            event.completed({ allowEvent: false });
                        }
                    }
                );
            } else {
                // Error reading body
                event.completed({ allowEvent: false });
            }
        }
    );
}

// Register the function
Office.actions.associate("insertPexipMeeting", insertPexipMeeting);