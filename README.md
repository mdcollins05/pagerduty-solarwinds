# PagerDuty for SolarWinds
This integration uses a VBScript that triggers and resolves incidents in PagerDuty when an alert is triggered or resets in SolarWinds. With this method, SolarWinds alerts go out in the form of a text file written to a queue directory (which contains data in PagerDuty's [Events API JSON format](https://v2.developer.pagerduty.com/docs/events-api)), and the VBScript sends the alert files in the queue directory to PagerDuty.

## Compatibility
This method works for SolarWinds products using the Orion 11.5+ backend, such as NPM and SAM.

## Usage
There are example alert triggers and resets for SolarWinds NPM and SAM in the **Alerts** directory. Importing these will start logging alerts to `C:\PagerDuty\Queue`.

For a full walkthrough on setting up the integration, please see the [SolarWinds Integration Guide](https://www.pagerduty.com/docs/guides/solarwinds-integration-guide/) on pagerduty.com.