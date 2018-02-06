$script:UnboundExtensionMethod = @"
using System;
using System.Collections;
using System.Management.Automation;
using System.Reflection;

namespace BAMCIS.PowerShell.Common
{
    public static class ExtensionMethods 
    {
        private static readonly BindingFlags Flags = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public;

        public static T GetUnboundParameterValue<T>(this PSCmdlet cmdlet, string paramName, int unnamedPosition = -1)
        {
            if (cmdlet != null)
            {
                // If paramName isn't found, value at unnamedPosition will be returned instead
                object Context = GetPropertyValue(cmdlet, "Context");
                object Processor = GetPropertyValue(Context, "CurrentCommandProcessor");
                object ParameterBinder = GetPropertyValue(Processor, "CmdletParameterBinderController");
                IEnumerable Args = GetPropertyValue(ParameterBinder, "UnboundArguments") as System.Collections.IEnumerable;

                if (Args != null)
                {
                    bool IsSwitch = typeof(SwitchParameter) == typeof(T);
                    string CurrentParameterName = String.Empty;
                    int i = 0;

                    foreach (object Arg in Args)
                    {

                        //Is the unbound argument associated with a parameter name
                        object IsParameterName = GetPropertyValue(Arg, "ParameterNameSpecified");

                        //The parameter name for the argument was specified
                        if (IsParameterName != null && true.Equals(IsParameterName))
                        {
                            string ParameterName = GetPropertyValue(Arg, "ParameterName") as string;
                            CurrentParameterName = ParameterName;

                            //If it's a switch parameter, there won't be a value following it, so just return a present switch
                            if (IsSwitch && String.Equals(CurrentParameterName, paramName, StringComparison.OrdinalIgnoreCase))
                            {
                                return (T)(object)new SwitchParameter(true);
                            }

                            //Since we have a current parameter name, the next value in Args should be the value supplied
                            //to the argument, so we can head on to the next iteration, this skips the remaining code below
                            //and starts the next item in the foreach loop
                            continue;
                        }

                        //We assume the previous iteration identified a parameter name, so this must be its
                        //value
                        object ParameterValue = GetPropertyValue(Arg, "ArgumentValue");

                        //If the value we have grabbed had a parameter name specified,
                        //let's check to see if it's the desired parameter
                        if (CurrentParameterName != String.Empty)
                        {
                            //If the parameter name currently being assessed is equal to the provided param
                            //name, then return the value of the param
                            if (CurrentParameterName.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                            {
                                return ConvertParameter<T>(ParameterValue);
                            }
                            else
                            {
                                //Since this wasn't the parameter name we were looking for, clear it out
                                CurrentParameterName = String.Empty;
                            }
                        }
                        //Otherwise there wasn't a parameter name, so the argument must have been supplied positionally,
                        //check if the current index is the position whose value we want
                        //Since positional parameters have to be specified first, this will be evaluated and increment until
                        //we run out of parameters or find a parameter with a name/value
                        else if (i++ == unnamedPosition)
                        {
                            //Just return the parameter value if the position matches what was specified
                            return ConvertParameter<T>(ParameterValue);
                        }
                    }
                }

                return default(T);
            }
            else
            {
                throw new ArgumentNullException("cmdlet", "The PSCmdlet cannot be null.");
            }
        }

        private static object GetPropertyValue(object instance, string fieldName)
        {
            // any access of a null object returns null. 
            if (instance == null || String.IsNullOrEmpty(fieldName))
            {
                return null;
            }

            try
            {
                PropertyInfo PropInfo = instance.GetType().GetProperty(fieldName, Flags);
            
                if (PropInfo != null)
                {
                    try
                    {
                        return PropInfo.GetValue(instance, null);
                    }
                    catch (Exception) { }
                }

                // maybe it's a field
                FieldInfo FInfo = instance.GetType().GetField(fieldName, Flags);

                if (FInfo != null)
                {
                    try
                    {
                        return FInfo.GetValue(instance);
                    }
                    catch { }
                }
            }
            catch (Exception) { }

            // no match, return null.
            return null;
        }
    
        private static T ConvertParameter<T>(this object value)
        {
            if (value == null || object.Equals(value, default(T)))
            {
                return default(T);
            }

            PSObject PSObj = value as PSObject;

            if (PSObj != null)
            {
                return PSObj.BaseObject.ConvertParameter<T>();
            }

            if (value is T)
            {
                if (typeof(T) == typeof(string))
                {
                    //Remove quotes from string values taken from the command line
                    // value = value.ToString().Trim('"').Trim('\'');
                }
                return (T)value;
            }

            var constructorInfo = typeof(T).GetConstructor(new[] { value.GetType() });

            if (constructorInfo != null)
            {
                return (T)constructorInfo.Invoke(new[] { value });
            }

            try
            {
                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch (Exception)
            {
                return default(T);
            }
        }    
    }
}
"@

#region Dynamic Parameters

Function Get-PropertyValue {
	<#
		.SYNOPSIS
			Attempts to get the value of a property on an object.

		.DESCRIPTION
			The cmdlet uses reflection to get the value of a property on the provided object. If the property does not exist, the cmdlet returns null.

		.PARAMETER InputObject
			The object instance to get the property value of.

		.PARAMETER Name
			The name of the object property or field to retrieve the value of.
		
		.EXAMPLE
			Get-PropertyValue -InputObject (New-Object -TypeName System.IO.FileInfo("c:\pagefile.sys")) -FieldName FullName

			This cmdlet returns the value "c:\pagefile.sys"

		.INPUTS
			System.Object

		.OUTPUTS
			System.Object
		
		.NOTES
            AUTHOR: Michael Haken
			LAST UPDATE: 2/6/2018

	#>
    [CmdletBinding()]
	[OutputType([System.Object])]
    Param(
        [Parameter(ValueFromPipeline = $true, Mandatory = $true, Position = 0)]
        [System.Object]$InputObject,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name
    )

    Begin {
        [System.Reflection.BindingFlags]$BindingFlags = @([System.Reflection.BindingFlags]::Instance, [System.Reflection.BindingFlags]::NonPublic, [System.Reflection.BindingFlags]::Public)
    }

    Process {
        if ($InputObject -eq $null -or [System.String]::IsNullOrEmpty($Name))
        {
            Write-Output -InputObject $null
        }

        [System.Reflection.PropertyInfo]$PropertyInfo = $InputObject.GetType().GetProperty($Name, $BindingFlags)
    
        if ($PropertyInfo -ne $null)
        {
            try {
				Write-Output -InputObject $PropertyInfo.GetValue($InputObject, $null)
            }
            catch [Exception] {
				Write-Verbose -Message "[ERROR] : $($_.Exception.Message)"
                Write-Output -InputObject $null
            }
        }
		# Maybe the property is a field
        else
        {
            [System.Reflection.FieldInfo]$FieldInfo = $InputObject.GetType().GetField($Name, $BindingFlags)

            if ($FieldInfo -ne $null)
            {
                try {
                    Write-Output -InputObject $FieldInfo.GetValue($InputObject, $null)
                }
                catch [Exception] {
					Write-Verbose -Message "[ERROR] : $($_.Exception.Message)"
                    Write-Output -InputObject $null
                }
            }
            else {
				# The name wasn't a property or field
                Write-Output -InputObject $null
            }
        }
    }

    End {
    }
}

Function Get-UnboundParameterValue {
	<#
		.SYNOPSIS
			Gets the value of an unbound dynamic parameter from an array of unbound parameters.

		.DESCRIPTION
			This cmdlet gets the value of a specified dynamic parameter name or positional parameter from the enumerated unbound dynamic parameters of a PowerShell cmdlet.

		.PARAMETER UnboundArgs
			The unbound arguments from a PowerShell cmdlet.

		.PARAMETER ParameterName
			The name of the parameter to get the value of.

		.PARAMETER Type
			The type of the parameter value.

		.PARAMETER Position 
			The position of the parameter to get the value of. Use this if the syntax 'New-Cmdlet -Parameter "Value"' was NOT used and instead 'New-Cmdlet "Value"' was used instead.

		.EXAMPLE
			DynamicParam {
				...
			
				[System.Reflection.BindingFlags]$BindingFlags = @([System.Reflection.BindingFlags]::Instance, [System.Reflection.BindingFlags]::NonPublic, [System.Reflection.BindingFlags]::Public)
				$Context = Get-PropertyValue -InputObject $PSCmdlet -Name "Context"
			
				# Can't use Get-PropertyValue fpr CurrentCommandProcessor because it returns itself as the current command processor
				$CurrentCommandProcessor = $Context.GetType().GetProperty("CurrentCommandProcessor", $BindingFlags).GetValue($Context)
				$ParameterBinder = Get-PropertyValue -InputObject $CurrentCommandProcessor -Name "CmdletParameterBinderController"
				$UnboundArgs = Get-PropertyValue -InputObject $ParameterBinder -Name "UnboundArguments"

				[System.String]$Target = (Get-UnboundParameterValue -UnboundArgs $UnboundArgs -ParameterName "Target" -Type ([System.String])) -as [System.String]

				...
			}

			This example enumerates the unbound arguments inside the dynamic parameter section of a PowerShell cmdlets. It supplies those arguments to the Get-UnboundParameterValue cmdlet looking
			for the value of the "Target" parameter. If the target parameter has been defined at the command line, the $Target variable will receive its value, otherwise null is returned.

		.INPUTS
			System.Object[]

		.OUTPUTS
			System.Object

		.NOTES
            AUTHOR: Michael Haken
			LAST UPDATE: 8/23/2017
	#>
	[CmdletBinding()]
	[OutputType([System.Object])]
	Param(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
		[System.Object[]]$UnboundArgs,

		[Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [System.String]$ParameterName,

        [Parameter(Mandatory = $true)]
        [System.Type]$Type,

        [Parameter(ParameterSetName = "Position", Mandatory = $true)]
        [System.Int32]$Position = -1
	)

	Begin {
	}

	Process {

		if ($UnboundArgs -ne $null)
        {
            [System.Boolean]$IsSwitch = [System.Management.Automation.SwitchParameter] -eq $Type

            [System.Int32]$i = 0

            foreach ($Item in $UnboundArgs | Where-Object {$_ -ne $null})
            {
				# Is the unbound argument associated with a parameter name
                $IsParameterName = Get-PropertyValue $Item -Name "ParameterNameSpecified"

				# The parameter name for the argument was specified
                if ($IsParameterName -ne $null -and $true.Equals($IsParameterName))
                {
                    [System.String]$CurrentParameterName = Get-PropertyValue $Item -Name "ParameterName"

					# If it's a switch parameter that was requested, there won't be a value following it, so just return a present switch
                    if ($IsSwitch -and [System.String]::Equals($CurrentParameterName, $ParameterName, [System.StringComparison]::OrdinalIgnoreCase))
                    {
						# Use return to stop execution
                        return (New-Object -TypeName System.Management.Automation.SwitchParameter($true))
                    }

					# Since we have a current parameter name, the next value in UnboundArgs should be the value supplied to the argument
					# so continue will start the next iteration in the foreach and skip the below code
                    continue
                }
                
				# We assume the previous iteration identified a parameter name, so this must be its value
                $ParamValue = Get-PropertyValue $Item -Name "ArgumentValue"

				if ($Type -eq [System.String])
				{
					$ParamValue = $ParamValue.Trim("`"").Trim("'")
				}

				# If the value we have grabbed had a parameter name specified, 
				# let's check to see if it's the desired parameter
                if (-not [System.String]::IsNullOrEmpty($CurrentParameterName))
                {
					# If the parameter name currently being assessed is equal to the provided param name, then return the value of the param
                    if ($CurrentParameterName.Equals($ParameterName, [System.StringComparison]::OrdinalIgnoreCase))
                    {
                        return $ParamValue 
                    }
                    else
                    {
						# Since this wasn't the parameter name we were looking for, clear it out
                        $CurrentParameterName = [System.String]::Empty
                    }
                }
				# Otherwise there wasn't a parameter name, so the argument must have been supplied positionally,
				# check if the current index is the position whose value we want.
				# Since positional parameters have to be specified frst, this will be evaluated and increment until
				# we run out of parameters or find a parameter with a name/value
                elseif ($i++ -eq $Position) 
				{
					return $ParamValue
                }
            }
        }
        else
        {
            Write-Output -InputObject $null
        }
	}

	End {
	}
}

Function Import-UnboundParameterCode {
	<#
		.SYNOPSIS
			Imports the .NET code to inspect unbound dynamic parameters in a PowerShell cmdlet DynamicParam section.

		.DESCRIPTION
			The cmdlet performs and Add-Type to import the code. It can also pass through the type you need to then invoke the unbound parameter checking.

		.PARAMETER PassThru
			Passes the static type to the pipeline.

		.EXAMPLE
			DynamicParam {
			...

				$Type = Import-UnboundParameterCode -PassThru
				$Type.GetMethod("GetUnboundParameterValue").MakeGenericMethod([System.String]).Invoke($Type, @($PSCmdlet, "Target", -1))

			...
			}
			
			This example imports the .NET code inside the DyanmicParam section of a PowerShell cmdlet. Then it uses the passed static class to call the 
			generic GetUnboundParameterValue method looking for the "Target" parameter. That parameter is a dynamic parameter added earlier in the DynamicParam section.

		.INPUTS
			None

		.OUTPUTS
			None or System.Type

		.NOTES
            AUTHOR: Michael Haken
			LAST UPDATE: 2/6/2018
			
	#>
	[CmdletBinding()]
	[OutputType([System.Type])]
	Param(
		[Parameter()]
		[Switch]$PassThru
	)

	Begin {
	}

	Process {
		if (-not ([System.Management.Automation.PSTypeName]"BAMCIS.PowerShell.Common.ExtensionMethods").Type) {
            Add-Type -TypeDefinition $script:UnboundExtensionMethod
			Write-Verbose -Message "Type BAMCIS.PowerShell.Common.ExtensionMethods successfully added."
        }
		else {
			Write-Verbose -Message "Type BAMCIS.PowerShell.Common.ExtensionMethods already added."
		}

		if ($PassThru) {
			Write-Output -InputObject ([BAMCIS.PowerShell.Common.ExtensionMethods])
		}
	}

	End {
	}
}

Function New-DynamicParameter {
	<#
		.SYNOPSIS
			Expedites creating PowerShell cmdlet dynamic parameters.

		.DESCRIPTION
			This cmdlet facilitates the easy creation of dynamic parameters.

		.PARAMETER Name
			The name of the parameter.

		.PARAMETER Type
			The type of the parameter, this defaults to System.String.

		.PARAMETER Mandatory
			Indicates whether the parameter is required when the cmdlet or function is run.

		.PARAMETER ParameterSets
			The name of the parameter sets to which this parameter belongs. This defaults to __AllParameterSets.

		.PARAMETER Position
			The position of the parameter in the command-line string.

		.PARAMETER ValueFromPipeline
			Indicates whether the parameter can take values from incoming pipeline objects.

		.PARAMETER ValueFromPipelineByPropertyName
			Indicates that the parameter can take values from a property of the incoming pipeline object that has the same name as this parameter. For example, if the name of the cmdlet or function parameter is userName, the parameter can take values from the userName property of incoming objects.

		.PARAMETER ValueFromRemainingArguments
			Indicates whether the cmdlet parameter accepts all the remaining command-line arguments that are associated with this parameter.

		.PARAMETER HelpMessage
			A short description of the parameter.

		.PARAMETER DontShow
			Indicates that this parameter should not be shown to the user in this like intellisense. This is primarily to be used in functions that are implementing the logic for dynamic keywords.

		.PARAMETER Alias
			Declares a alternative namea for the parameter.

		.PARAMETER ValidateNotNull
			Validates that the argument of an optional parameter is not null.

		.PARAMETER ValidateNotNullOrEmpty
			Validates that the argument of an optional parameter is not null, an empty string, or an empty collection.

		.PARAMETER AllowEmptyString
			Allows Empty strings.

		.PARAMETER AllowNull
			Allows null values.

		.PARAMETER AllowEmptyCollection
			Allows empty collections.

		.PARAMETER ValidateScript
			Defines an attribute that uses a script to validate a parameter of any Windows PowerShell function.

		.PARAMETER ValidateSet
			Defines an attribute that uses a set of values to validate a cmdlet parameter argument.

		.PARAMETER ValidateRange
			Defines an attribute that uses minimum and maximum values to validate a cmdlet parameter argument.

		.PARAMETER ValidateCount
			Defines an attribute that uses maximum and minimum limits to validate the number of arguments that a cmdlet parameter accepts.

		.PARAMETER ValidateLength
			Defines an attribute that uses minimum and maximum limits to validate the number of characters in a cmdlet parameter argument.

		.PARAMETER ValidatePattern
			Defines an attribute that uses a regular expression to validate the character pattern of a cmdlet parameter argument.

		.PARAMETER RuntimeParameterDictionary
			The dictionary to add the new parameter to. If one is not provided, a new dictionary is created and returned to the pipeline.
		
		.EXAMPLE
			DynamicParam {
				...

				$RuntimeParameterDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary

				New-DynamicParameter -Name "Numbers" -ValidateSet @(1, 2, 3) -Type [System.Int32] -Mandatory -RuntimeParameterDictionary $RuntimeParameterDictionary | Out-Null

				...

				return $RuntimeParameterDictionary

			}

			A new parameter named "Numbers" is added to the cmdlet. The parameter is mandatory and must be 1, 2, or 3. The dictionary sent in is modified and does not need to be received. 

		.EXAMPLE
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

			The example creates an array of two hashtables. These hashtables are converted into PSObjects so they can match the parameters by property name, then new dynamic parameters are created. All of the 
			parameters are fed to New-DynamicParameter which returns a single new RuntimeParameterDictionary to the pipeline, which is returned from the DynamicParam section.

		.INPUTS
			System.Management.Automation.PSObject

		.OUTPUTS
			System.Management.Automation.RuntimeDefinedParameterDictionary

		.NOTES
            AUTHOR: Michael Haken
			LAST UPDATE: 2/6/2018
	#>
	[CmdletBinding()]
	[OutputType([System.Management.Automation.RuntimeDefinedParameterDictionary])]
	Param(
		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$Name,

		# These parameters are part of the standard ParameterAttribute

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNull()]
		[System.Type]$Type = [System.String],

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$Mandatory,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateCount(1, [System.Int32]::MaxValue)]
		[System.String[]]$ParameterSets = @("__AllParameterSets"),

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[System.Int32]$Position = [System.Int32]::MinValue,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$ValueFromPipeline,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$ValueFromPipelineByPropertyName,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$ValueFromRemainingArguments,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$HelpMessage,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$DontShow,

		# These parameters are each their own attribute

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[System.String[]]$Alias = @(),

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$ValidateNotNull,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$ValidateNotNullOrEmpty,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$AllowEmptyString,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$AllowNull,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[Switch]$AllowEmptyCollection,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[System.Management.Automation.ScriptBlock]$ValidateScript,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNull()]
		[System.String[]]$ValidateSet = @(),

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateCount(2,2)]
		[System.Int32[]]$ValidateRange = $null,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateCount(2,2)]
		[System.Int32[]]$ValidateCount = $null,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateCount(2,2)]
		[System.Int32[]]$ValidateLength = $null,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$ValidatePattern = $null,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNull()]
		[System.Management.Automation.RuntimeDefinedParameterDictionary]$RuntimeParameterDictionary = $null
	)

	Begin {
		if ($RuntimeParameterDictionary -eq $null) {
			$RuntimeParameterDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
		}
	}

	Process {
		# Create the collection of attributes
		$AttributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]
			
		foreach ($Set in $ParameterSets)
		{
			# Create and set the parameter's attributes
			$ParameterAttribute = New-Object -TypeName System.Management.Automation.PARAMETERAttribute

			if (-not [System.String]::IsNullOrEmpty($Set))
			{
				$ParameterAttribute.ParameterSetName = $Set
			}

			if ($Position -ne $null)
			{
				$ParameterAttribute.Position = $Position
			}

			if ($Mandatory)
			{
				$ParameterAttribute.Mandatory = $true
			}

			if ($ValueFromPipeline)
			{
				$ParameterAttribute.ValueFromPipeline = $true
			}

			if ($ValueFromPipelineByPropertyName)
			{
				$ParameterAttribute.ValueFromPipelineByPropertyName = $true
			}

			if ($ValueFromRemainingArguments)
			{
				$ParameterAttribute.ValueFromRemainingArguments = $true
			}

			if (-not [System.String]::IsNullOrEmpty($HelpMessage))
			{
				$ParameterAttribute.HelpMessage = $HelpMessage
			}

			if ($DontShow)
			{
				$ParameterAttribute.DontShow = $true
			}

			$AttributeCollection.Add($ParameterAttribute)
		}

		if ($Alias.Length -gt 0)
		{
			$AliasAttribute = New-Object -TypeName System.Management.Automation.AliasAttribute($Alias)
			$AttributeCollection.Add($AliasAttribute)
		}

		if ($ValidateSet.Length -gt 0)
		{
			$ValidateSetAttribute = New-Object -TypeName System.Management.Automation.ValidateSetAttribute($ValidateSet)
			$AttributeCollection.Add($ValidateSetAttribute)
		}

		if ($ValidateScript -ne $null) 
		{
			$ValidateScriptAttribute = New-Object -TypeName System.Management.Automation.ValidateScriptAttribute($ValidateScript)
			$AttributeCollection.Add($ValidateScriptAttribute)
		}

		if ($ValidateCount -ne $null -and $ValidateCount.Length -eq 2)
		{
			$ValidateCountAttribute = New-Object -TypeName System.Management.Automation.ValidateCountAttribute($ValidateCount[0], $ValidateCount[1])
			$AttributeCollection.Add($ValidateCountAttribute)
		}

		if ($ValidateLength -ne $null -and $ValidateLength -eq 2)
		{
			$ValidateLengthAttribute = New-Object -TypeName System.Management.Automation.ValidateLengthAttribute($ValidateLength[0], $ValidateLength[1])
			$AttributeCollection.Add($ValidateLengthAttribute)
		}

		if (-not [System.String]::IsNullOrEmpty($ValidatePattern))
		{
			$ValidatePatternAttribute = New-Object -TypeName System.Management.Automation.ValidatePatternAttribute($ValidatePattern)
			$AttributeCollection.Add($ValidatePatternAttribute)
		}

		if ($ValidateRange -ne $null -and $ValidateRange.Length -eq 2)
		{
			$ValidateRangeAttribute = New-Object -TypeName System.Management.Automation.ValidateRangeAttribute($ValidateRange)
			$AttributeCollection.Add($ValidateRangeAttribute)
		}

		if ($ValidateNotNull)
		{
			$NotNullAttribute = New-Object -TypeName System.Management.Automation.ValidateNotNullAttribute
			$AttributeCollection.Add($NotNullAttribute)
		}

		if ($ValidateNotNullOrEmpty)
		{
			$NotNullOrEmptyAttribute = New-Object -TypeName System.Management.Automation.ValidateNotNullOrEmptyAttribute
			$AttributeCollection.Add($NotNullOrEmptyAttribute)
		}

		if ($AllowEmptyString)
		{
			$AllowEmptyStringAttribute = New-Object -TypeName System.Management.Automation.AllowEmptyStringAttribute
			$AttributeCollection.Add($AllowEmptyStringAttribute)
		}

		if ($AllowEmptyCollection)
		{
			$AllowEmptyCollectionAttribute = New-Object -TypeName System.Management.Automation.AllowEmptyCollectionAttribute
			$AttributeCollection.Add($AllowEmptyCollectionAttribute)
		}

		if ($AllowNull)
		{
			$AllowNullAttribute = New-Object -TypeName System.Management.Automation.AllowNullAttribute
			$AttributeCollection.Add($AllowNullAttribute)
		}

		if (-not $RuntimeParameterDictionary.ContainsKey($Name))
		{
			$RuntimeParameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter($Name, $Type, $AttributeCollection)
			$RuntimeParameterDictionary.Add($Name, $RuntimeParameter)
		}
		else
		{
			foreach ($Attr in $AttributeCollection.GetEnumerator())
            {
                if (-not $RuntimeParameterDictionary.$Name.Attributes.Contains($Attr))
                {
                    $RuntimeParameterDictionary.$Name.Attributes.Add($Attr)
                }
            }
		}
	}

	End {
		Write-Output -InputObject $RuntimeParameterDictionary
	}
}

#endregion