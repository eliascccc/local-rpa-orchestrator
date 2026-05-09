# place in main.py directory

import re
from main import JobCandidate, JobResult, JobName, HandoverFile

class QtyChangeHandler:
    '''Example of a personal inbox email automation for demo purpose.'''
   
    job_name: JobName = "qty_adjust"
    
    def __init__(self, logger) -> None:
        self.logger = logger


    def can_handle(self, candidate: JobCandidate) -> bool:
        subject = str(candidate.email_subject).strip().lower()
        return self.job_name in subject


    def precheck_and_build_payload(self, candidate: JobCandidate) -> JobResult:
        ''' sanity-check on given data (including ERP check if possible)'''
        email_body = candidate.email_body
        assert email_body is not None # to satisfy pylance

        # get relevant info for qty_adjust, eg:
        order_number_match = re.search(r"order_number:\s*(.+)", email_body)
        order_number = order_number_match.group(1) if order_number_match else None

        order_qty_match = re.search(r"order_qty:\s*(.+)", email_body)
        order_qty = order_qty_match.group(1) if order_qty_match else None

        material_available_match = re.search(r"material_available:\s*(.+)", email_body)
        material_available = material_available_match.group(1) if material_available_match else None

        error_message = ""
        if order_number is None:
            error_message += "missing order_number. "
        if order_qty is None:
            error_message += "missing order_qty. "
        if material_available is None:
            error_message += "missing material_available. "

        if error_message:
            return JobResult(is_success=False, error_message=error_message.strip())

        # and for any attachments, eg:
        attachments = candidate.parsed_source_data.get("attachments", [])
        #for attachment in attachments:
        #    print(attachment.get("filename"))

        rpatool_payload = {
            "order_number": order_number,
            "order_qty": order_qty,
            "target_order_qty": material_available,
            "attachments": attachments,
        }

        return JobResult(is_success=True, rpatool_payload=rpatool_payload)
    

    def verify_result(self, handover_file: HandoverFile) -> JobResult:
        # placeholder for implementation
        return JobResult(is_success=True)



def build_custom_personal_mail_handlers(logger) -> dict:
    return {
        "qty_adjust": QtyChangeHandler(logger),
    }
