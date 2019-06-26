# - Create a firewall rule on this system.

# - Choose a name for the new rule.

$DisplayName = ""

# - Define direction of rule. Can be "Inbound" or "Outbound".

$Direction = ""

# - Define ports for this rule. More than one port can be entered, seperated by comma's.

$PortNumbers = ""

# - Define the protocol. Can be TCP or UDP.

$Protocol = ""

# - Action can be either "Allow" or "Block".

$Action = ""

# - Command for executing this action.

New-NetFirewallRule -DisplayName "$DisplayName" -Direction $Direction -LocalPort $PortNumbers -Protocol $Protocol -Action $Action