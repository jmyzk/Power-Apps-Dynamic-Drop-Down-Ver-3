// ====================================================================
// btnCommitFieldConfiguration.OnSelect - FIXED GUID VERSION
// Purpose: Save field configuration to new TargetSheetColumnDefinition list
// ====================================================================

// 1. VALIDATION
If(
    IsBlank(varCurrentFieldColumnId) || IsBlank(varFormID),
    Notify(
        "Missing field or form information.",
        NotificationType.Error
    );
    Exit()
);

// 2. GET CURRENT FIELD DATA
Set(
    varCurrentField,
    LookUp(
        colUnifiedFields,
        targetColumnId = varCurrentFieldColumnId
    )
);

// 3. VALIDATE FIELD EXISTS
If(
    IsBlank(varCurrentField),
    Notify(
        "Field not found in collection.",
        NotificationType.Error
    );
    Exit()
);

// 4. SAVE/UPDATE FIELD TO FLAT LIST
// First, check if record exists (delegation-friendly with proper GUID handling)
Set(
    varExistingRecord,
    LookUp(
        TargetSheetColumnDefinition,
        FormID = Text(GUID(varFormID)) && TargetColumnID = varCurrentFieldColumnId
    )
);

// Update existing or create new record
If(
    IsBlank(varExistingRecord),
    // CREATE NEW RECORD
    Patch(
        TargetSheetColumnDefinition,
        Defaults(TargetSheetColumnDefinition),
        {
            Title: "Primary Field " & varCurrentField.displayOrder & " - " & varCurrentField.targetColumnTitle,
            RecordID: GUID(),
            FormID: Text(GUID(varFormID)), // Ensure proper GUID text format
            FieldType: varCurrentField.section,
            DisplayPosition: varCurrentField.displayOrder,
            
            // Target Column Information
            TargetColumnID: varCurrentField.targetColumnId,
            TargetColumnTitle: varCurrentField.targetColumnTitle,
            TargetColumnType: varCurrentField.targetColumnType,
            TargetColumnIndex: varCurrentField.targetColumnIndex,
            SymbolSubtype: "",
            
            // Field Behavior Settings
            IsRequired: varCurrentField.isRequired,
            IsActive: true,
            IsConditional: false,
            ConditionalRulesJSON: "",
            
            // Cascade Specific Fields (Not applicable for primary)
            SourceColumnID: "",
            SourceColumnTitle: "",
            SourceColumnType: "",
            
            // Option Configuration
            OptionSourceType: varOptionsSource,
            Options:  If(
                varCurrentField.targetColumnType in ["PICKLIST", "MULTI_PICKLIST"],
                JSON(varPreviewOptions),
                ""
            ),
            ContactOptions: If(
                varCurrentField.targetColumnType in ["CONTACT_LIST", "MULTI_CONTACT_LIST"],
                JSON(varPreviewOptions),
                ""
            ),
            
            // External Sheet Configuration
            ExternalSheetID: If(varOptionsSource = "external-source", varConfirmedExternalSheetID, ""),
            ExternalSheetName: If(varOptionsSource = "external-source", varConfirmedSheetName, ""),
            ExternalColumnID: If(varOptionsSource = "external-source", Text(varConfirmedExternalColumn.id), ""),
            ExternalColumnTitle: If(varOptionsSource = "external-source", varConfirmedExternalColumn.title, ""),
            ExternalColumnType: If(varOptionsSource = "external-source", varConfirmedExternalColumn.type, ""),
            
            // Metadata
            ConfigurationComplete: true,
            LastUpdated: Now(),
            CreatedDate: Now()
        }
    ),
    // UPDATE EXISTING RECORD
    Patch(
        TargetSheetColumnDefinition,
        varExistingRecord,
        {
            // Update only the fields that can change
            DisplayPosition: varCurrentField.displayOrder,
            IsRequired: varCurrentField.isRequired,
            
            // Option Configuration
            OptionSourceType: varOptionsSource,
            Options:If(
                varCurrentField.targetColumnType in ["PICKLIST", "MULTI_PICKLIST"],
                JSON(varPreviewOptions),
                ""
            ),
            ContactOptions: If(
                varCurrentField.targetColumnType in ["CONTACT_LIST", "MULTI_CONTACT_LIST"],
                JSON(varPreviewOptions),
                ""
            ),
            
            // External Sheet Configuration
            ExternalSheetID: If(varOptionsSource = "external-source", varConfirmedExternalSheetID, ""),
            ExternalSheetName: If(varOptionsSource = "external-source", varConfirmedSheetName, ""),
            ExternalColumnID: If(varOptionsSource = "external-source", Text(varConfirmedExternalColumn.id), ""),
            ExternalColumnTitle: If(varOptionsSource = "external-source", varConfirmedExternalColumn.title, ""),
            ExternalColumnType: If(varOptionsSource = "external-source", varConfirmedExternalColumn.type, ""),
            
            // Metadata
            ConfigurationComplete: true,
            LastUpdated: Now()
        }
    )
);





// 5. UPDATE FORM METADATA (No JSON - just counts)
// removed this part as TargetSheetColumnDefinition share point list has the information. (JSON part > the list)


// 6. UPDATE EXTERNAL SHEETS REGISTRY (only if external source selected)
If(
    varOptionsSource = "external-source" && varExternalSelectionConfirmed,
    Patch(
        'External Option Sheets',
        LookUp('External Option Sheets', SheetID = varConfirmedExternalSheetID),
        {
            UsageCount: LookUp('External Option Sheets', SheetID = varConfirmedExternalSheetID, UsageCount) + 1,
            LastSynced: Now()
        }
    )
);


// 7. UPDATE COLLECTION STATUS
Patch(
    colColumnOptionConfiguration,
    LookUp(colColumnOptionConfiguration, columnId = varCurrentFieldColumnId),
    {optionSourceSummary: "Configured"}
);

// 8. UPDATE LOCAL COLLECTION
Patch(
    colUnifiedFields,
    LookUp(colUnifiedFields, targetColumnId = varCurrentFieldColumnId),
    {
        hasExternalSource: varOptionsSource = "external-source" && varExternalSelectionConfirmed,
        optionSourceType: varOptionsSource,
        isConfigured: true,
        configurationComplete: true,
        lastModified: Now()
    }
);


// 9. RESET CONFIGURATION STATE
Set(varExternalSelectionConfirmed, false);
Set(varCurrentFieldColumnId, "");
Set(varShowOptionConfiguration    , false);

// 10. SUCCESS NOTIFICATION
With(
    {
        fieldTitle: varCurrentField.targetColumnTitle,
        sourceType: Switch(
            varOptionsSource,
            "external-source", "External Source (" & varConfirmedSheetName & ")",
            "target-combined", "Target Sheet + Column Definition",
            "Column Definition Only"
        )
    },
    Notify(
        "Field '" & fieldTitle & "' configured with " & sourceType & " - saved to flat structure!",
        NotificationType.Success,
        3000
    )
);

// 11. TRIGGER UI REFRESH
Set(varGlobalRefresh, !varGlobalRefresh);


/// the following is not in claude 

// 12. Update colColumnOptionConfiguration to show "Configured" status
Set(
    varThisColumnId,
    LookUp(colColumnOptionConfiguration, columnId = varCurrentFieldColumnId).columnId

);
ClearCollect(
    colText,
    Filter(
        colColumnOptionConfiguration,
        columnId = varThisColumnId
    )
);
Patch(
    colColumnOptionConfiguration,
    LookUp(
        colColumnOptionConfiguration,
        columnId = varThisColumnId
    ),
    {optionSourceSummary: "Configured"}
);

// clear the checkbox selection


// Close the panel
Set(varShowOptionConfiguration    , false);

// Reset option source
// Set(varOptionsSource, "none");

// Reset enable variables (updated for 3-checkbox structure)
Set(varEnableTargetColumn, true);
Set(varEnableTargetCombined, true);
Set(varEnableExternalSource, true);

// Reset checkboxes (updated for 3-checkbox structure)
Reset(chkTargetColumn);
Reset(chkTargetCombined);
Reset(chkExternalSource);

// ✅ NEW: Reset preview-related variables
Set(varPreviewGenerated, false);
Set(varHasUncommittedChanges, false);
Set(varCommittingChanges, false);
Set(varPreviewOptions, Table());

// Reset Power Automate related variables
Set(varUseUpdatedOptions, false);
ClearCollect(colUpdatedOptions, {});

// ✅ NEW: Reset external source variables (for future implementation)
Set(varShowSheetSourcePanel, false);

// ✅ NEW: Clear temporary collections
ClearCollect(colPreview_ColumnOptions, []);
ClearCollect(colPreview_SheetData, []);
ClearCollect(colPreview_CombinedValues, []);
ClearCollect(colPreview_SplitValues, []);





