"""Tamil (India) text catalog for user-facing strings."""

TEXTS = {
    # Startup / main
    "app.already_running": "பயன்பாடு ஏற்கனவே இயங்கிக்கொண்டிருக்கிறது.",
    "app.unexpected_error": "எதிர்பாராத பிழை ஏற்பட்டது. பதிவு கோப்பை சரிபார்க்கவும்.",
    "app.startup.workbook_secret_missing_frozen": "பணிப்புத்தகம் ரகசியம் கிடைக்கவில்லை. மென்பொருளை மீண்டும் நிறுவவும் அல்லது ஆதரவைக் தொடர்புகொள்ளவும்.",
    "app.startup.workbook_secret_missing_dev": "பணிப்புத்தகம் ரகசிய சேமிப்பகத்தை உள்ளூரில் நீக்கி பயன்பாட்டை மீண்டும் தொடங்கவும்.",
    "app.main_window_title": "ஃபோகஸ்",
    "splash.starting": "பயன்பாடு தொடங்குகிறது...",
    "splash.loading_main_window": "முதன்மை சாளரம் ஏற்றப்படுகிறது...",
    # Main window
    "toolbar.navigation": "வழிசெலுத்தல்",
    "status.ready": "தயார்",
    "common.dropzone.placeholder": "கோப்புகளை சேர்க்க Drag and Drop செய்யவும், அல்லது Ctrl + O அழுத்தவும், அல்லது ஒருமுறை கிளிக் செய்யவும்",
    "module.placeholder": "{title} தொகுதி இன்னும் செயல்படுத்தப்படவில்லை.",
    "module.instructor": "பயிற்றுநர்",
    "module.coordinator_short": "ஒருங்கிணைப்பாளர்",
    "module.po_analysis": "PO பகுப்பாய்வு",
    "module.co_analysis": "CO பகுப்பாய்வு",
    "module.load_failed_title": "தொகுதி ஏற்றப் பிழை",
    "module.load_failed_body": "'{module}' தொகுதியை ஏற்ற முடியவில்லை.\n\nவிவரம்: {error}",
    "module.load_failed_status": "தொகுதியை ஏற்ற முடியவில்லை: {module}.",
    "nav.help": "உதவி",
    "nav.about": "பற்றி",
    # Settings module
    "language.switcher.button": "மொழி: {language}",
    "language.switcher.applied_status": "மொழி {language} ஆக அமைக்கப்பட்டது.",
    # About
    "about.version": "பதிப்பு {version}",
    "about.subtitle": "ஃப்ரேம்வொர்க் ஃபார் அவுட்கம் கம்ப்யூடேஷன் அண்ட் யூனிபிகேஷன் சிஸ்டம்",
    "about.description": (
        "{app_name} என்பது கோர்ஸ் அவுட்கம் (CO) அடைவை கணக்கிடவும், "
        "நேரடி மற்றும் மறைமுக மதிப்பீடுகளை அடிப்படையாகக் கொண்டு "
        "கட்டமைக்கப்பட்ட முடிவு பகுப்பாய்வை செய்யவும் வடிவமைக்கப்பட்ட மென்பொருள் கருவி."
    ),
    "about.institution": "கலசலிங்கம் ஆராய்ச்சி மற்றும் கல்வி கழகத்தில் (KARE) பின்வரும் பங்களிப்பாளர்களால் உருவாக்கப்பட்டது:",
    "about.copyright": "(c) {year} KARE. அனைத்து உரிமைகளும் பாதுகாக்கப்பட்டவை.",
    "about.contributors.none": "பங்களிப்பாளர் தகவல் கிடைக்கவில்லை.",
    "about.repository.link_label": "கிட்ஹப் ரெப்போசிட்டரி",
    "about.meta_line_html": '{copyright} | <a href="{url}">{link_label}</a>',
    # Help module
    "help.doc_missing_title": "உதவி ஆவணம் இல்லை",
    "help.doc_missing_body": "உதவி PDF கிடைக்கவில்லை:\\n{path}",
    "help.doc_error_title": "உதவி ஆவணப் பிழை",
    "help.doc_error_body": "உதவி PDF ஐ ஏற்ற முடியவில்லை.",
    "help.download_pdf": "PDF பதிவிறக்கு",
    "help.open_default_viewer": "இயல்புநிலை பார்வியில் திற",
    "help.missing_file_title": "கோப்பு இல்லை",
    "help.save_title": "CO அடைவு ஆவணத்தை சேமி",
    "help.save_default_name": "CO_Calculation_Document.pdf",
    "help.save_filter_pdf": "PDF கோப்புகள் (*.pdf)",
    "help.save_failed_title": "சேமித்தல் தோல்வி",
    "help.save_success_title": "சேமிக்கப்பட்டது",
    "help.open_failed_title": "திறத்தல் தோல்வி",
    "help.open_failed_body": "இயல்புநிலை பார்வியில் உதவி PDF ஐ திறக்க முடியவில்லை.",
    "help.open_success_title": "திறக்கப்பட்டது",
    "help.open_success_body": "உதவி PDF இயல்புநிலை பார்வியில் திறக்கப்பட்டது.",
    "help.status.doc_loaded": "உதவி ஆவணம் ஏற்றப்பட்டது.",
    "help.status.doc_missing": "உதவி ஆவணம் இல்லை.",
    "help.status.doc_error": "உதவி ஆவணத்தை ஏற்ற முடியவில்லை.",
    "help.status.file_missing": "உதவி PDF கிடைக்கவில்லை.",
    "help.status.save_success": "உதவி PDF வெற்றிகரமாக சேமிக்கப்பட்டது.",
    "help.status.save_failed": "உதவி PDF சேமிக்கும் போது பிழை ஏற்பட்டது.",
    "help.status.open_success": "உதவி PDF இயல்புநிலை பார்வியில் திறக்கப்பட்டது.",
    "help.status.open_failed": "உதவி PDF திறக்கும் போது பிழை ஏற்பட்டது.",
    "outputs.title": "உருவாக்கப்பட்ட வெளியீடுகள்",
    "activity.log.title": "செயற்பாட்டு பதிவு",
    "activity.log.ready": "செயற்பாட்டு பதிவு தொடங்கப்பட்டது.",
    "outputs.none_generated": "வெளியீடுகள் எதுவும் உருவாக்கப்படவில்லை.",
    "outputs.open_file": "கோப்பை திற",
    "outputs.open_folder": "கோப்புறையை திற",
    "outputs.open_failed": "தேர்ந்தெடுக்கப்பட்ட பாதையை திறக்க முடியவில்லை.",
    # Instructor module
    "instructor.workflow_title": "CO பணிச்சுற்று",
    "instructor.links.title": "உருவாக்கப்பட்ட வெளியீடுகள்",
    "instructor.links.course_details_generated": "உருவாக்கப்பட்ட பாட விவர வார்ப்புரு",
    "instructor.links.marks_template_generated": "உருவாக்கப்பட்ட மதிப்பெண் வார்ப்புரு",
    "instructor.links.open_file": "கோப்பை திற",
    "instructor.links.open_folder": "கோப்புறையை திற",
    "instructor.links.open_failed": "தேர்ந்தெடுக்கப்பட்ட பாதையை திறக்க முடியவில்லை.",
    "instructor.action.download_course_template": "பாட வார்ப்புருவை பதிவிறக்கு",
    "instructor.action.download_course_template_link_html": '<a href="{href}">{text}</a>',
    "instructor.action.generate_marks_template": "மதிப்பெண் வார்ப்புருவை தயாரி",
    "instructor.msg.success_title": "வெற்றி",
    "instructor.msg.validation_title": "சரிபார்ப்பு பிழை",
    "instructor.msg.error_title": "தோல்வி",
    "instructor.msg.failed_to_do": "{action} செய்ய முடியவில்லை.",
    "instructor.dialog.filter.excel": "Excel கோப்புகள் (*.xlsx)",
    "instructor.dialog.filter.excel_open": "Excel கோப்புகள் (*.xlsx *.xlsm *.xls)",
    "instructor.dialog.course_template.save_title": "பாட விவர வார்ப்புருவை சேமிக்க",
    "instructor.dialog.course_template.default_name": "course_details_template.xlsx",
    "instructor.dialog.course_details.select_title": "பாட விவர கோப்பை தேர்வு செய்ய",
    "instructor.dialog.marks_template.save_title": "மதிப்பெண் வார்ப்புருவை சேமிக்க",
    "instructor.dialog.marks_template.default_name": "marks_template.xlsx",
    "instructor.status.template_download_path_selected": "வார்ப்புரு பதிவிறக்கப் பாதை தேர்வு செய்யப்பட்டது.",
    "instructor.status.course_details_validated": "பாட விவரங்கள் பதிவேற்றப்பட்டு சரிபார்க்கப்பட்டன. இப்போது மதிப்பெண் வார்ப்புருவை தயாரிக்கலாம்.",
    "instructor.status.marks_template_generated": "மதிப்பெண் வார்ப்புரு உருவாக்கப்பட்டது.",
    "instructor.status.course_details_replaced": "பாட விவரங்கள் மாற்றப்பட்டன. அடுத்த படிகளை மீண்டும் செய்யவும்.",
    "instructor.status.operation_cancelled": "செயல் ரத்து செய்யப்பட்டது.",
    "instructor.status.course_details_drop_browse_requested": "இழுத்து-விடும் விட்ஜெட்டில் பாட விவர கோப்பு தேர்வு கோரப்பட்டது.",
    "instructor.status.course_details_drop_files_dropped": "பாட விவர விட்ஜெட்டில் {count} கோப்பு(கள்) விடப்பட்டது.",
    "instructor.status.course_details_drop_files_changed": "பாட விவர விட்ஜெட்டில் இப்போது {count} கோப்பு(கள்) உள்ளன.",
    "instructor.status.course_details_drop_files_rejected": "பாட விவர விட்ஜெட்டில் விடப்பட்ட {count} கோப்பு(கள்) நிராகரிக்கப்பட்டன.",
    "instructor.status.course_details_validation_progress": "செல்லுபடியாகிய கோப்புகள் சரிபார்க்கப்பட்டவை: {valid}/{total}.",
    "instructor.status.marks_template_generation_progress": "மதிப்பெண் வார்ப்புருக்கள் செயலாக்கப்பட்டவை: {processed}/{total}.",
    "instructor.status.marks_template_per_file_failures": "மதிப்பெண்-வார்ப்புரு கோப்பு-தோறும் தோல்விகள்: {details}",
    "instructor.drop.summary": "கோப்புகள்: {count}",
    "instructor.toast.course_details_validation_summary": "சரிபார்ப்பு முடிந்தது: {valid} செல்லுபடியாகும், {invalid} செல்லுபடியாகாதது, {mismatched} தவறான வார்ப்புரு, {duplicates} நகல் உள்ளீடு.",
    "instructor.toast.marks_template_generation_summary": "மதிப்பெண் வார்ப்புரு உருவாக்கம் முடிந்தது: செயலாக்கப்பட்டது {processed}/{total}, உருவாக்கப்பட்டது {generated}, தோல்வி {failed}, தவிர்க்கப்பட்டது {skipped}.",
    "instructor.log.process.generate_course_details_template": "பாட விவர வார்ப்புரு உருவாக்கப்படுகிறது",
    "instructor.log.process.validate_course_details_workbook": "பதிவேற்றப்பட்ட பாட விவர பணிப்புத்தகம் சரிபார்க்கப்படுகிறது",
    "instructor.log.process.generate_marks_template": "மதிப்பெண் வார்ப்புரு உருவாக்கப்படுகிறது",
    "instructor.validation.file_issue_line": "{file}: [{code}] {reason}",
    "instructor.validation.course_details_missing": "மதிப்பெண் வார்ப்புரு உருவாக்குவதற்கு முன் குறைந்தது ஒரு செல்லுபடியாகும் பாட விவர பணிப்புத்தகத்தை பதிவேற்றவும்.",
    "validation.course_details.duplicate_path": "Duplicate file path skipped: {workbook}.",
    "validation.course_details.duplicate_section": "Duplicate section skipped for this cohort (section '{section}') in file: {workbook}.",
    "validation.course_details.cohort_mismatch": "Cohort mismatch in {workbook}. These fields must match the first valid file: {fields}.",
    "validation.course_details.unexpected_rejection": "File was skipped due to an unexpected validation failure: {workbook}.",
    "coordinator.thresholds.l1.label": "L1 வரம்பு:",
    "coordinator.thresholds.l2.label": "L2 வரம்பு:",
    "coordinator.thresholds.l3.label": "L3 வரம்பு:",
    "coordinator.thresholds.invalid_rule": "வரம்பு விதி மீறப்பட்டது: 0 < L1 < L2 < L3 < 100 என்பதைப் பின்பற்றவும்.",
    "coordinator.co_attainment.description": "குறிப்பிட்ட நிலை மற்றும் அதற்கு மேலுள்ள மாணவர்கள் % அடிப்படையில் அடைந்ததாக கருதப்படும்.",
    "coordinator.co_attainment.percent.label": "CO AT%:",
    "coordinator.co_attainment.level.label": "CO AT நிலை:",
    "coordinator.co_attainment.invalid_percent": "தவறான CO AT%: மதிப்பு 0 முதல் 100 வரை இருக்க வேண்டும்.",
    "coordinator.dialog.title": "Excel கோப்புகளை தேர்வு செய்ய",
    "coordinator.links.downloaded_output": "பதிவிறக்கப்பட்ட வெளியீடு",
    "coordinator.summary": "கோப்புகள்: {count}",
    "coordinator.file.remove_fallback": "நீக்கு",
    "coordinator.file.remove_tooltip": "கோப்பை நீக்கு",
    "coordinator.clear_all": "அனைத்தையும் நீக்கு",
    "coordinator.status.added": "{added} கோப்பு(கள்) சேர்க்கப்பட்டது. மொத்தம்: {total}.",
    "coordinator.duplicate.title": "கோப்பு ஏற்கனவே உள்ளது",
    "coordinator.duplicate.body": "பட்டியலில் ஏற்கனவே உள்ளதால் {count} நகல் கோப்பு(கள்) தவிர்க்கப்பட்டன.",
    "coordinator.invalid_final_report.title": "தவறான இறுதி CO அறிக்கை",
    "coordinator.invalid_final_report.body": "செல்லுபடியான இறுதி CO report பணிப்புத்தகங்கள் மட்டுமே அனுமதிக்கப்படும்.\n\n{files}",
    "coordinator.invalid_final_report.details_prefix": "விவரங்கள்:",
    "coordinator.invalid_final_report.detail_line": "{file}: {reason}",
    "coordinator.status.ignored": (
        "{count} கோப்பு(கள்) புறக்கணிக்கப்பட்டன (ஆதரிக்காத வகை, இல்லை, நகல், "
        "அல்லது தவறான final report workbook)."
    ),
    "coordinator.status.removed": "{count} கோப்பு(கள்) நீக்கப்பட்டன.",
    "coordinator.status.cleared": "{count} கோப்பு(கள்) அழிக்கப்பட்டன.",
    "coordinator.status.processing_started": "தேர்ந்தெடுக்கப்பட்ட கோப்புகள் செயலாக்கப்படுகின்றன...",
    "coordinator.status.queued": "நடப்பு செயலாக்கம் நடைபெறுவதால் {count} கோப்பு(கள்) வரிசைப்படுத்தப்பட்டன.",
    "coordinator.status.operation_cancelled": "ஒருங்கிணைப்பாளர் செயலாக்கம் ரத்து செய்யப்பட்டது.",
    "coordinator.status.processing_failed": "ஒருங்கிணைப்பாளர் செயலாக்கம் தோல்வியடைந்தது.",
    "coordinator.status.processing_completed": "ஒருங்கிணைப்பாளர் செயலாக்கம் முடிந்தது.",
    "coordinator.status.calculate_completed": "CO அடைவு கணக்கீடு நிறைவடைந்து வெளியீட்டு பணிப்புத்தகம் உருவாக்கப்பட்டது.",
    "coordinator.regno_dedup.title": "நகல் பதிவு எண்கள் நீக்கப்பட்டன",
    "coordinator.regno_dedup.body": "CO அடைவு உருவாக்கும்போது {count} நகல் பதிவு எண் பதிவுகள் நீக்கப்பட்டன.",
    "coordinator.regno_dedup.log_body": "{count} நகல் பதிவு எண் பதிவுகள் நீக்கப்பட்டன:\n{details}",
    "coordinator.regno_dedup.log_detail": "பதிவு எண்: {reg_no} | பணித்தாள்: {worksheet} | பணிப்புத்தகம்: {workbook}",
    "coordinator.regno_dedup.log_detail_unavailable": "நகல் பதிவு விவரங்கள் கிடைக்கவில்லை.",
    "coordinator.join_drop.body": "சில வரிகள் Direct அல்லது Indirect தாள்களில் மட்டும் இருந்ததால் புறக்கணிக்கப்பட்டன. நீக்கப்பட்ட வரிகள்: {count}. விவரங்களுக்கு செயல் பதிவைப் பார்க்கவும்.",
    "co_analysis.status.duplicate_reg_numbers": "நகல் பதிவு எண்கள் காரணமாக {count} கோப்பு(கள்) நிராகரிக்கப்பட்டன.",
    "co_analysis.status.co_count_mismatch": "மொத்த விளைவு எண்ணிக்கை பொருந்தாததால் {count} கோப்பு(கள்) நிராகரிக்கப்பட்டன.",
    "co_analysis.status.validation_warnings": "{count} சரிபார்ப்பு எச்சரிக்கை(கள்) அறிவிக்கப்பட்டன.",
    "co_analysis.status.rejected_code_with_context": "{file} நிராகரிக்கப்பட்டது: {code} ({context})",
    "co_analysis.dialog.select_files": "ஆய்விற்கான கோப்புகளை தேர்வு செய்யவும்",
    "co_analysis.validation.co_count_mismatch_body": "தேர்ந்தெடுத்த தொகுப்புடன் மொத்த விளைவு எண்ணிக்கை வேறுபட்டதால் {count} கோப்பு(கள்) நிராகரிக்கப்பட்டன.",
    "co_analysis.validation.anomaly_warnings_body": "பதிவேற்றிய கோப்புகளில் {count} சரிபார்ப்பு அசாதாரண எச்சரிக்கை(கள்) அறிக்கையிடப்பட்டன.",
    "co_analysis.validation.rejection_breakdown_title": "CO Analysis கோப்பு நிராகரிப்புகள்",
    "co_analysis.validation.rejection_breakdown_body": (
        "காரணப்படி நிராகரிக்கப்பட்ட கோப்புகள்:\\n"
        "ஆதரிக்காதது/இல்லை: {unsupported_or_missing}\\n"
        "தவறான பணிப்புத்தகம்: {invalid_workbook}\\n"
        "சிஸ்டம் ஹாஷ்/அடையாள தோல்வி: {invalid_hash}\\n"
        "நிரப்பப்படாத மதிப்பெண்கள்: {marks_unfilled}\\n"
        "லேஅவுட்/மனிபெஸ்ட் தோல்வி: {layout_manifest}\\n"
        "வார்ப்புரு பொருந்தாமை: {template_mismatch}\\n"
        "தவறான மார்க் மதிப்பு/வடிவம்: {mark_value}\\n"
        "பிற சரிபார்ப்பு தோல்விகள்: {invalid_other}\\n"
        "நகல்கள்: {duplicates}\\n"
        "நகல் பதிவு எண்கள்: {duplicate_reg}\\n"
        "மொத்த விளைவு எண்ணிக்கை பொருந்தாமை: {co_count_mismatch}"
    ),
    "co_analysis.status.ignored_summary": (
        "{count} கோப்பு(கள்) புறக்கணிக்கப்பட்டன. "
        "{part1}{sep1_2}{part2}{sep2_3}{part3}{sep3_4}{part4}{sep4_5}{part5}{sep5_6}{part6}"
        "{sep6_7}{part7}{sep7_8}{part8}{sep8_9}{part9}{sep9_10}{part10}{sep10_11}{part11}{sep11_12}{part12}"
    ),
    "co_analysis.status.ignored_total": "{count} கோப்பு(கள்) புறக்கணிக்கப்பட்டன.",
    "co_analysis.status.ignored_reason.unsupported_or_missing": "ஆதரிக்காதது/இல்லை={count}",
    "co_analysis.status.ignored_reason.invalid_workbook": "தவறான பணிப்புத்தகம்={count}",
    "co_analysis.status.ignored_reason.invalid_hash": "ஹாஷ்={count}",
    "co_analysis.status.ignored_reason.marks_unfilled": "நிரப்பப்படாத மதிப்பெண்கள்={count}",
    "co_analysis.status.ignored_reason.layout_manifest": "லேஅவுட்/மனிபெஸ்ட்={count}",
    "co_analysis.status.ignored_reason.template_mismatch": "வார்ப்புரு பொருந்தாமை={count}",
    "co_analysis.status.ignored_reason.mark_value": "தவறான மார்க் மதிப்பு={count}",
    "co_analysis.status.ignored_reason.invalid_other": "பிற சரிபார்ப்பு={count}",
    "co_analysis.status.ignored_reason.duplicates": "நகல்கள்={count}",
    "co_analysis.status.ignored_reason.duplicate_reg": "நகல் பதிவு எண்கள்={count}",
    "co_analysis.status.ignored_reason.co_count_mismatch": "மொத்த விளைவு எண்ணிக்கை பொருந்தாமை={count}",
}








