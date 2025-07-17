class CPKNotFoundError(FileNotFoundError):
    """Raised if a CPK datasheet file could not be found"""
    def __init__(self, message):
        super().__init__(message)

class ReportExistsError(FileExistsError):
    """
    Custom error for existing files
    Raised if a report already exists for a specific combination of 
    model_code + (furnace_code/extrusion_batch_code)
    """
    def __init__(self, message):
        super().__init__(message)

class NonConformantError(Exception):
    """Raised when functional performance does not meet required standards."""
    def __init__(self, message="Non-conformant to required specifications"):
        self.message = message
        super().__init__(self.message)