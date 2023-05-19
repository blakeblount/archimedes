import json

class Config:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialize_config()
        return cls._instance

    def _initialize_config(self):
        with open("config.json", "r") as config_file:
            self._config_data = json.load(config_file)

    # Getter methods for file paths
    def get_pattern_work_order_template_file(self):
        return self._config_data["file_paths"]["pattern_work_order_template_file"]

    def get_pattern_work_order_word_folder(self):
        return self._config_data["file_paths"]["pattern_work_order_Word_folder"]

    def get_pattern_work_order_pdf_folder(self):
        return self._config_data["file_paths"]["pattern_work_order_PDF_folder"]

    def get_pattern_receipt_template_file(self):
        return self._config_data["file_paths"]["pattern_receipt_template_file"]

    def get_pattern_receipt_word_folder(self):
        return self._config_data["file_paths"]["pattern_receipt_Word_folder"]

    def get_pattern_receipt_pdf_folder(self):
        return self._config_data["file_paths"]["pattern_receipt_PDF_folder"]

    def get_drawings_folder(self):
        return self._config_data["file_paths"]["drawings_folder"]

    def get_engineering_log_book_file(self):
        return self._config_data["file_paths"]["engineering_log_book_file"]

    # Getter methods for Syteline settings
    def get_server(self):
        return self._config_data["syteline_settings"]["server"]

    def get_database(self):
        return self._config_data["syteline_settings"]["database"]

    def get_user_id(self):
        return self._config_data["syteline_settings"]["user_id"]

    def get_user_pwd(self):
        return self._config_data["syteline_settings"]["user_pwd"]

    # Getter methods for personal email settings
    def get_smtp_server(self):
        return self._config_data["personal_email_settings"]["smtp_server"]

    def get_smtp_port(self):
        return self._config_data["personal_email_settings"]["smtp_port"]

    def get_personal_email_address(self):
        return self._config_data["personal_email_settings"]["personal_email_address"]

    def get_personal_email_token(self):
        return self._config_data["personal_email_settings"]["personal_email_token"]

    # Getter methods for Jira settings
    def get_jira_base_url(self):
        return self._config_data["jira_settings"]["jira_base_url"]

    def get_jira_username(self):
        return self._config_data["jira_settings"]["jira_username"]

    def get_jira_api_token(self):
        return self._config_data["jira_settings"]["jira_api_token"]

    # Getter methods for outgoing email addresses
    def get_purchasing_email_address(self):
        return self._config_data["outgoing_email_addresses"]["purchasing_email_address"]

    def get_intertool_email_address(self):
        return self._config_data["outgoing_email_addresses"]["intertool_email_address"]

    def get_pro_pattern_email_address(self):
        return self._config_data["outgoing_email_addresses"]["pro_pattern_email_address"]
