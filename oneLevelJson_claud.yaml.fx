// ====================================================================
// btnCommitFieldConfiguration.OnSelect - SEMI-FLAT SHAREPOINT VERSION
// Purpose: Save field configuration to both existing JSON and new flat lists
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

// 2. SET VARIABLES (Using new pattern from migration plan)
Set(
    varCurrentField,
    LookUp(
        colPrimaryFields,
        targetColumnId = varCurrentFieldColumnId
    )
);

Set(
    varHasExternalSource,
    varOptionsSource = "external-source"
);

Set(
    varExternalSourceType,
    If(varOptionsSource = "external-source", "EXTERNAL_SHEET", "")
);

Set(
    varOptionSourceConfig,
    Switch(
        varOptionsSource,
        "external-source",
        JSON([{
            type: "EXTERNAL_SHEET",
            sheetId: varConfirmedExternalSheetID,
            sheetName: varConfirmedSheetName,
            columnId: varConfirmedExternalColumn.id,
            columnTitle: varConfirmedExternalColumn.title,
            columnType: varConfirmedExternalColumn.type,
            previewCount: CountRows(varPreviewOptions),
            configuredDate: Text(Now(), "yyyy-mm-ddThh:mm:ssZ")
        }]),
        "target-combined",
        JSON([{
            type: "TARGET_COMBINED",
            useColumnDefinition: true,
            useTargetSheetData: true,
            configuredDate: Text(Now(), "yyyy-mm-ddThh:mm:ssZ")
        }]),
        JSON([{
            type: "COLUMN_DEFINITION",
            useColumnDefinition: true,
            configuredDate: Text(Now(), "yyyy-mm-ddThh:mm:ssZ")
        }])
    )
);

// 3. GET CURRENT SHAREPOINT DATA
With(
    {
        currentRecord: LookUp(
            'Form Definition Admin',
            FormID = Text(varFormID)
        ),
        currentConfig: ParseJSON(
            LookUp(
                'Form Definition Admin',
                FormID = Text(varFormID),
                FormConfiguration
            )
        )
    },
    
    // 4. BUILD FIELD CONFIGURATION for JSON (backward compatibility)
    With(
        {
            finalFieldConfiguration: {
                fieldId: varCurrentField.fieldId,
                targetColumnId: varCurrentField.targetColumnId,
                targetColumnTitle: varCurrentField.targetColumnTitle,
                targetColumnType: varCurrentField.targetColumnType,
                targetColumnIndex: varCurrentField.targetColumnIndex,
                controlType: varCurrentField.controlType,
                section: "primary",
                isRequired: varCurrentField.isRequired,
                allowMultiSelect: varCurrentField.allowMultiSelect,
                displayOrder: varCurrentField.displayOrder,
                optionSourceType: varOptionsSource,
                hasExternalSource: varHasExternalSource,
                externalSourceType: varExternalSourceType,
                externalSheetId: If(varOptionsSource = "external-source", varConfirmedExternalSheetID, ""),
                externalSheetName: If(varOptionsSource = "external-source", varConfirmedSheetName, ""),
                externalColumnId: If(varOptionsSource = "external-source", Text(varConfirmedExternalColumn.id), ""),
                externalColumnTitle: If(varOptionsSource = "external-source", varConfirmedExternalColumn.title, ""),
                externalColumnType: If(varOptionsSource = "external-source", varConfirmedExternalColumn.type, "")
            },
            
            currentPrimaryFields: If(
                IsBlank(currentConfig.primaryFields) || CountRows(Table(currentConfig.primaryFields)) = 0,
                [],
                Table(currentConfig.primaryFields)
            )
        },
        
        // 5. UPDATE EXISTING JSON SHAREPOINT RECORD (backward compatibility)
        ClearCollect(colPrimaryField, varCurrentField);
        Collect(colPrimaryField, finalFieldConfiguration);
        Set(
            varUpdatedFormConfigJSON,
            JSON(
                {
                    formMetadata: {
                        version: "1.2",
                        stage: Text(currentConfig.formMetadata.stage),
                        created: Text(currentConfig.formMetadata.created),
                        lastModified: Text(Now(), "yyyy-mm-ddThh:mm:ssZ")
                    },
                    sheetConfiguration: currentConfig.sheetConfiguration,
                    cascadeFields: currentConfig.cascadeFields,
                    primaryFields: colPrimaryFields,
                    secondaryFields: currentConfig.secondaryFields
                }
            )
        );
        
        // Update Form Definition Admin (existing system)
        Patch(
            'Form Definition Admin',
            LookUp('Form Definition Admin', FormID = Text(varFormID)),
            {
                FormConfiguration: varUpdatedFormConfigJSON,
                ModifiedDate: Now(),
                PrimaryCount: CountRows(colPrimaryFields),
                UsesExternalOptions: varHasExternalSource,
                ExternalOptionCount: If(varHasExternalSource, 1, 0)
            }
        )
    )
);

// 6. SAVE TO NEW FLAT SHAREPOINT LISTS

// 6a. Save to PrimaryFields (new flat table)
Patch(PrimaryFields,
    Defaults(PrimaryFields),
    {
        PrimaryFieldID: GUID(),
        FormID: varFormID,
        FieldId: varCurrentField.fieldId,
        DisplayOrder: varCurrentField.displayOrder,
        TargetColumnId: varCurrentField.targetColumnId,
        TargetColumnTitle: varCurrentField.targetColumnTitle,
        TargetColumnType: varCurrentField.targetColumnType,
        TargetColumnIndex: varCurrentField.targetColumnIndex,
        ControlType: varCurrentField.controlType,
        Section: "primary",
        IsRequired: varCurrentField.isRequired,
        AllowMultiSelect: varCurrentField.allowMultiSelect,
        OptionSourceType: varOptionsSource,
        HasExternalSource: varHasExternalSource,
        ExternalSourceType: varExternalSourceType,
        ExternalSheetId: If(varOptionsSource = "external-source", varConfirmedExternalSheetID, ""),
        ExternalSheetName: If(varOptionsSource = "external-source", varConfirmedSheetName, ""),
        ExternalColumnId: If(varOptionsSource = "external-source", Text(varConfirmedExternalColumn.id), ""),
        ExternalColumnTitle: If(varOptionsSource = "external-source", varConfirmedExternalColumn.title, ""),
        ExternalColumnType: If(varOptionsSource = "external-source", varConfirmedExternalColumn.type, ""),
        ConfigurationComplete: true,
        LastModified: Now()
    }
);

// 6b. Save option source configuration to OptionSourceConfigurations (new flat table)
Patch(OptionSourceConfigurations,
    Defaults(OptionSourceConfigurations),
    {
        ConfigID: GUID(),
        FormID: varFormID,
        FieldID: varCurrentField.fieldId,
        ConfigType: Switch(
            varOptionsSource,
            "external-source", "EXTERNAL_SHEET",
            "target-combined", "TARGET_COMBINED", 
            "COLUMN_DEFINITION"
        ),
        UseColumnDefinition: varOptionsSource in ["column-definition", "target-combined"],
        UseTargetSheetData: varOptionsSource = "target-combined",
        PreviewCount: If(varOptionsSource = "external-source", CountRows(varPreviewOptions), 0),
        ConfiguredDate: Now(),
        ConfigurationJSON: varOptionSourceConfig
    }
);

// 7. UPDATE EXTERNAL SHEETS REGISTRY (only if external source selected)
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

// 8. UPDATE LOCAL COLLECTION for UI display
Patch(
    colPrimaryFields,
    LookUp(colPrimaryFields, targetColumnId = varCurrentFieldColumnId),
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
Set(varShowConfigPanel, false);

// 10. SUCCESS NOTIFICATION
With(
    {
        fieldTitle: LookUp(colPrimaryFields, targetColumnId = varCurrentFieldColumnId, targetColumnTitle),
        sourceType: Switch(
            varOptionsSource,
            "external-source", "External Source (" & varConfirmedSheetName & ")",
            "target-combined", "Target Sheet + Column Definition",
            "Column Definition Only"
        )
    },
    Notify(
        "Field '" & fieldTitle & "' configured with " & sourceType & " - saved to both systems!",
        NotificationType.Success,
        3000
    )
);

// 11. TRIGGER UI REFRESH
Set(varGlobalRefresh, !varGlobalRefresh);
