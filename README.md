# BAMCIS Dynamic Parameters

## Usage

### Example 1

    DynamicParam {
	    ...
		$RuntimeParameterDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary

		New-DynamicParameter -Name "Numbers" -ValidateSet @(1, 2, 3) -Type [System.Int32] -Mandatory -RuntimeParameterDictionary $RuntimeParameterDictionary | Out-Null

		...

		return $RuntimeParameterDictionary
	}

This creates a new dynamic parameter named Numbers that has a validation set of 1, 2, and 3, and is mandatory.

### Example 2

    DynamicParam {
	    ...

		$Params = @(
			@{
				"Name" = "Numbers";
				"ValidateSet" = @(1, 2, 3);
				"Type" = [System.Int32]
			},
			@{
				"Name" = "FirstName";
				"Type" = [System.String];
				"Mandatory" = $true;
				"ParameterSets" = @("Names")
			}
		)

		$Params | ForEach-Object {
			New-Object PSObject -Property $_ 
		} | New-DynamicParameter
	}

The example creates an array of two hashtables. These hashtables are converted into PSObjects so they can match the parameters by property name, then new dynamic parameters are created. All of the parameters are fed to New-DynamicParameter which returns a single new RuntimeParameterDictionary to the pipeline, which is returned from the DynamicParam section.

## Revision History

### 1.0.0.0
Initial Release. This module has been separated from the BAMCIS.Common module to provide a lighter weight module that is more reusable across other modules.