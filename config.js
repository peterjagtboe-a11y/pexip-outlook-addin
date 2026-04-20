// Pexip Add-in Configuration
// Edit these values for each customer deployment

var PexipConfig = {
    // Default VMR domain (e.g., 'pexip.vc', 'meet.framskak.com', 'customer.infinity.com')
    vmrDomain: 'pexip.vc',
    
    // Optional: Pre-fill VMR username (leave empty to auto-detect from email)
    // The add-in will automatically use the user's email prefix (e.g., john.doe@company.com → john.doe)
    // Only set this if you want to override auto-detection
    defaultVmrUsername: '',
    
    // Company branding
    companyName: 'Pexip',
    
    // Meeting details text (customize the labels if needed)
    labels: {
        header: 'PEXIP MEETING DETAILS',
        webBrowser: 'Join from web browser:',
        vcEndpoint: 'Join from VC endpoint:',
        pexipApp: 'Join from Pexip App:'
    }
};