// ============================================================
// btnCommitFieldSettings.OnSelect
// Purpose: Save FIELD SETTINGS (required, default, validation, logic)
//          into TargetSheetColumnDefinition for the selected field
// ============================================================

// 0) Guards
If(
    IsBlank(varSelectedPrimaryField) || IsBlank(varFormID),
    Notify("Missing selected field or FormID.", NotificationType.Error); Exit()
);

// 1) Normalize IDs (delegation-friendly)
Set(varFormIdText, Text(GUID(varFormID)));
Set(varFieldIdText, Text(varSelectedPrimaryField.targetColumnId));

// 2) Pull current settings + rule + conditions from local collections
Set(varSettings,     LookUp(colFieldSettings,     TargetColumnId = varFieldIdText));
Set(varRuleRow,      LookUp(colLogicRules,        TargetColumnId = varFieldIdText));
ClearCollect(colMyConds, Filter(colLogicConditions, TargetColumnId = varFieldIdText && RuleID = Coalesce(varRuleRow.RuleID, "")));

// 3) Build JSON payloads
Set(varValidationJSON, JSON(varSettings)); // IsRequired, DefaultValue, InputMode, Min/Max, ValidationType, RegexPattern
Set(varRulesJSON,      JSON({ Rule: varRuleRow, Conditions: colMyConds }));

// 4) Find existing TSCD row (per Form + TargetColumn)
Set(
    varExistingTSCD,
    LookUp(
        TargetSheetColumnDefinition,
        FormID = varFormIdText && TargetColumnID = varFieldIdText
    )
);

// 5) Create or Update the TSCD row with FIELD SETTINGS
If(
    IsBlank(varExistingTSCD),

    /* CREATE */
    Patch(
        TargetSheetColumnDefinition,
        Defaults(TargetSheetColumnDefinition),
        {
            // Keys / identity
            RecordID: GUID(),
            FormID: varFormIdText,

            // Titles (from selected field)
            Title: "Field - " & varSelectedPrimaryField.targetColumnTitle,
            FieldType: varSelectedPrimaryField.section,                    // PRIMARY / SECONDARY / CASCADE
            DisplayPosition: varSelectedPrimaryField.displayOrder,

            // Target column
            TargetColumnID: varFieldIdText,
            TargetColumnTitle: varSelectedPrimaryField.targetColumnTitle,
            TargetColumnType: varSelectedPrimaryField.targetColumnType,
            TargetColumnIndex: varSelectedPrimaryField.targetColumnIndex,

            // FIELD SETTINGS (this panel)
            IsRequired: Coalesce(varSettings.IsRequired, false),
            DefaultValue: Coalesce(varSettings.DefaultValue, ""),
            IsConditional: !IsBlank(varRuleRow) && Coalesce(varRuleRow.Enabled, false),
            ConditionalRulesJSON: varRulesJSON,        // includes JoinMode/Action/Enabled + conditions
            ValidationRulesJSON: varValidationJSON,    // full field settings for validation

            // Metadata
            ConfigurationComplete: true,
            LastUpdated: Now(),
            CreatedDate: Now()
        }
    ),

    /* UPDATE */
    Patch(
        TargetSheetColumnDefinition,
        varExistingTSCD,
        {
            // Only fields that can change from this panel
            DisplayPosition: varSelectedPrimaryField.displayOrder,
            IsRequired: Coalesce(varSettings.IsRequired, false),
            DefaultValue: Coalesce(varSettings.DefaultValue, ""),
            IsConditional: !IsBlank(varRuleRow) && Coalesce(varRuleRow.Enabled, false),
            ConditionalRulesJSON: varRulesJSON,
            ValidationRulesJSON: varValidationJSON,
            LastUpdated: Now()
        }
    )
);

// 6) Local UI state updates (optional)
Patch(
    colUnifiedFields,
    LookUp(colUnifiedFields, Text(targetColumnId) = varFieldIdText),
    {
        isConfigured: true,
        configurationComplete: true,
        lastModified: Now()
    }
);

// 7) Close panel + notify
Set(varShowFieldConfiguration, false);
Notify(
    "Field settings saved for '" & varSelectedPrimaryField.targetColumnTitle & "'.",
    NotificationType.Success, 2500
);

// 8) Trigger UI refresh if you use it elsewhere
Set(varGlobalRefresh, !varGlobalRefresh);
