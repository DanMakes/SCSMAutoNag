<p align="center">
  <h1 align="center">AutoNag for SCSM</h1>
</p>

Regardless of the volume of Incidents a service desk deals with, a common task for employees is to follow up and stay on top of Incidents that haven't been modified in some period of time. This PowerShell script streamlines that process by:
- Evaluating Action Log entries
- When Action Log entries were left
- Who left Comments on the Action Log
- Sends HTML based emails to the Affected User and includes SCSM mailbox to be picked up by the Exchange Connector
- Resolves the Incident after several days of inactivity

This script borrows the Add-ActionLogEntry function from [SMLets Exchange Connector](https://github.com/AdhocAdam/smletsexchangeconnector) to manage the Action Log
