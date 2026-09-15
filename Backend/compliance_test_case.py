import os
import logging

class ComplianceTestPipeline:
    """
    This class introduces simulated architectural configurations 
    that should trigger compliance and governance penalties.
    """
    def __init__(self):
        # 1. Mock Static Credentials Configuration (Triggers Credentials Scanner Flags)
        self.api_key_storage = "MOCK_STATIC_DEVELOPMENT_KEY_XYZ123"
        self.database_url = "postgresql://mock_user:mock_password_plaintext@localhost:5432/hiremind_db"
        
        # 2. Storage Operational Parameters (Triggers Data Retention Flags)
        self.retain_all_logs = True
        self.enable_data_encryption = False # Simulates missing cryptographic controls
        self.output_directory = "./unencrypted_test_records/"

    def log_candidate_transaction(self, candidate_id, raw_input_string):
        """
        Simulates unvetted raw string logging to test input-handling analysis.
        """
        if not os.path.exists(self.output_directory):
            os.makedirs(self.output_directory)
            
        log_file = os.path.join(self.output_directory, "transaction_history.log")
        
        # High-risk trigger: Writing raw, un-sanitized user strings directly to file logs
        with open(log_file, "a") as f:
            f.write(f"ID: {candidate_id} - Payload: {raw_input_string}\n")
            
        print(f"[Compliance Test] Transaction logged for candidate {candidate_id}")
