// ====================================================================
// btnCommitFieldConfiguration.OnSelect - COMMIT FIELD CONFIG TO SHAREPOINT
// Purpose: Final commit of field configuration based on user's checkbox selection
// ====================================================================
// ====================================================================
// FIXED: btnCommitFieldConfiguration.OnSelect - Handle Empty primaryFields Array
// Purpose: Final commit of field configuration based on user's checkbox selection
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
// 2. GET CURRENT FIELD AND SET VARIABLES
Set(
    varCurrentField,
    LookUp(
        colPrimaryFields,
        targetColumnId = varCurrentFieldColumnId
    )
);

// 3. SET SOURCE-SPECIFIC CONFIGURATION VARIABLES
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
        {
            type: "EXTERNAL_SHEET",
            sheetId: varConfirmedExternalSheetID,
            sheetName: varConfirmedSheetName,
            columnId: varConfirmedExternalColumn.id,
            columnTitle: varConfirmedExternalColumn.title,
            columnType: varConfirmedExternalColumn.type,
            previewCount: CountRows(varPreviewOptions),
            configuredDate: Text(
                Now(),
                "yyyy-mm-ddThh:mm:ssZ"
            )
        },
        "target-combined",
        {
            type: "TARGET_COMBINED",
            useColumnDefinition: true,
            useTargetSheetData: true,
            configuredDate: Text(
                Now(),
                "yyyy-mm-ddThh:mm:ssZ"
            )
        },
        {
            type: "COLUMN_DEFINITION",
            useColumnDefinition: true,
            configuredDate: Text(
                Now(),
                "yyyy-mm-ddThh:mm:ssZ"
            )
        }
    )
);

// 4. GET SHAREPOINT DATA AND BUILD CONFIGURATION
With(
    {
        // Get current SharePoint record
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
    // 5. BUILD FIELD CONFIGURATION using pre-set variables
    With(
        {
            finalFieldConfiguration: {
                // === BASIC/TARGET SHEET PROPERTIES ===
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
                
                // === OPTION SOURCE PROPERTIES (GENERAL) ===
                optionSourceType: varOptionsSource,
                optionSourceConfig: varOptionSourceConfig,
                
                // === COLUMN DEFINITION ONLY ===
                // (No additional properties needed)
                
                // === TARGET COMBINED ===
                // (Properties handled in optionSourceConfig)
                
                // === EXTERNAL SHEET PROPERTIES (MOST COMPLEX) ===
                hasExternalSource: varHasExternalSource,
                externalSourceType: varExternalSourceType,
                externalSheetId: If(varOptionsSource = "external-source", varConfirmedExternalSheetID, ""),
                externalSheetName: If(varOptionsSource = "external-source", varConfirmedSheetName, ""),
                externalColumnId: If(varOptionsSource = "external-source", Text(varConfirmedExternalColumn.id), ""),
                externalColumnTitle: If(varOptionsSource = "external-source", varConfirmedExternalColumn.title, ""),
                externalColumnType: If(varOptionsSource = "external-source", varConfirmedExternalColumn.type, "")
            },
            // ✅ FIXED: Handle empty primaryFields array properly
            currentPrimaryFields: // Table(currentConfig.primaryFields)// Convert to table if exists
            If(
                IsBlank(currentConfig.primaryFields) || CountRows(Table(currentConfig.primaryFields)) = 0,
                [],// Empty array if primaryFields is null or empty
                Table(currentConfig.primaryFields)// Convert to table if exists
            )
        },
        // 6. UPDATE SHAREPOINT JSON CONFIGURATION
        ClearCollect(colPrimaryField, varCurrentField);
        Collect(colPrimaryField,finalFieldConfiguration);
        Set(
            varUpdatedFormConfigJSON,
            JSON(
                {
                    formMetadata: {
                        version: "1.2",
                        stage: Text(currentConfig.formMetadata.stage),
                        created: Text(currentConfig.formMetadata.created),
                        lastModified: Text(
                            Now(),
                            "yyyy-mm-ddThh:mm:ssZ"
                        )
                    },
                    sheetConfiguration: currentConfig.sheetConfiguration,
                    cascadeFields: currentConfig.cascadeFields,
                    // Update primaryFields array
                    // Simpler: Always append (for single-field commits)
                    primaryFields: colPrimaryFields, // [finalFieldConfiguration],
                    secondaryFields: currentConfig.secondaryFields
                }
            )
        )// End Set and JSON
    )// End With for section 5
);
// End With for section 4
// 7. UPDATE SHAREPOINT RECORD
With(
    {
        // Count external fields directly from local collection instead of parsing JSON
        externalFieldsCount: CountRows(
            Filter(
                colPrimaryFields,
                hasExternalSource = true
            )
        )
    },
    Patch(
        'Form Definition Admin',
        LookUp(
            'Form Definition Admin',
            FormID = Text(varFormID)
        ),
        {
            FormConfiguration: varUpdatedFormConfigJSON,
            ModifiedDate: Now(),
            PrimaryCount: CountRows(colPrimaryFields),
            UsesExternalOptions: externalFieldsCount > 0,
            ExternalOptionCount: externalFieldsCount
                // Simplified: Just store the current external sheet ID if one is being configured
        }
    )
);
// 8. UPDATE EXTERNAL SHEETS REGISTRY (only if external source selected)
If(
    varOptionsSource = "external-source" && varExternalSelectionConfirmed,
    Patch(
        'External Option Sheets',
        LookUp(
            'External Option Sheets',
            SheetID = varConfirmedExternalSheetID
        ),
        {
            UsageCount: LookUp(
                'External Option Sheets',
                SheetID = varConfirmedExternalSheetID,
                UsageCount
            ) + 1,
            LastSynced: Now()
        }
    )
);
// 9. UPDATE LOCAL COLLECTION for UI display
Patch(
    colPrimaryFields,
    LookUp(
        colPrimaryFields,
        targetColumnId = varCurrentFieldColumnId
    ),
    {
        hasExternalSource: varOptionsSource = "external-source" && varExternalSelectionConfirmed,
        optionSourceType: varOptionsSource,
        isConfigured: true,
        configurationComplete: true,
        lastModified: Now()
    }
);
// 10. RESET CONFIGURATION STATE
Set(
    varExternalSelectionConfirmed,
    false
);
Set(
    varCurrentFieldColumnId,
    ""
);
Set(
    varShowConfigPanel,
    false
);
// 11. SUCCESS NOTIFICATION
With(
    {
        fieldTitle: LookUp(
            colPrimaryFields,
            targetColumnId = varCurrentFieldColumnId,
            targetColumnTitle
        ),
        sourceType: Switch(
            varOptionsSource,
            "external-source",
            "External Source (" & varConfirmedSheetName & ")",
            "target-combined",
            "Target Sheet + Column Definition",
            "Column Definition Only"
        )
    },
    Notify(
        "Field '" & fieldTitle & "' configured with " & sourceType,
        NotificationType.Success,
        3000
    )
);
// 12. TRIGGER UI REFRESH
Set(
    varGlobalRefresh,
    !varGlobalRefresh
);