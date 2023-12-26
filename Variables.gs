var specialFormatTextForNo = `Issue Reported: "Commander not onboarded" on Register screen
Did it ever work: Yes
Activity at Time of Occurrence:{Space}
- Issue started{Space}
Steps Taken Before Call:
-{Space}
KB/Resource(s) utilized: KB 7789
Action/Capture Information:{Space}

contact person:{Space}
callback:{Space}

Troubleshooting:
- Site is on Base 53.28 > Cloud Connect feature is available{Space}
- Verified the Service ID: [] > matches from Help > Support{Space}
- Advised site to complete the signup for the Verifone C-site management and not to skip any of the steps > site understood{Space}
- Advised the site that the message is a notification for them to onboard their commander to C-site management and will only be removed after they have finished "Cloud Connect"{Space}
- Site will cb after they have finished signing up to request assistance for cloud connect{Space}
- Site provided email:{Space}
- Verifone c-site management signup email sent{Space}

[**]

Results: Site will cb after they have finished signing up to request assistance for cloud connect.

Worked With: Escalation Lead{Space}`;

var specialFormatTextForYes = `Issue Reported: "Commander not onboarded" on Register screen{Space}
Did it ever work: Yes
Activity at Time of Occurrence:
- Issue started{Space}
Steps Taken Before Call: 
-{Space}
KB/Resource(s) utilized: KB 7789
Action/Capture Information:{Space}

contact person:{Space}
callback:{Space}

Troubleshooting:
- Verified with the site that they have already signed up for the Verifone C-site management and has already creaded Hierarchy, Site, Roles and Users{Space}
- Site is on Base 53.28 > Cloud Connect feature is available{Space}
- Verified the Service ID: [] > matches from Help > Support{Space}
- Ran cat /home/rsd/config/service_id > service ID ={Space}
- Had site login to configuration manager > Clicked 'Cloud Connect' from Initial Setup > User code: [] > Clicked the link{Space} 
- Had site enter the user code in the text field > Submit{Space}
- Had site Login using Merchant Admin credentials > Clicked "Allow"{Space} 
- Had site navigate to Initial Setup > Cloud Connect > "Confirm Verification"{Space}


Results: Cloud connect has been successfully setup.{Space}

Worked With: Escalation Lead{Space}`;
