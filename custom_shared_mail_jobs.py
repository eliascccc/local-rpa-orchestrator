# place in main.py directory

import re
from main import JobCandidate, JobResult, JobName, HandoverFile

class PoAdjustHandler:
    '''Example of a shared inbox email automation for demo purpose.'''

    job_name: JobName = "po_adjust"

    def __init__(self, logger) -> None:
        self.logger = logger


    def can_handle(self, candidate: JobCandidate) -> bool:
        # Placeholder for mailbox-specific in scope rules, eg:
        return str(candidate.email_address) == "supplier1@example.com" and "Order confirmation" in str(candidate.email_subject)


    def precheck_and_build_payload(self, candidate: JobCandidate) -> JobResult:
        ''' sanity-check on given data (including ERP check if possible)'''
        email_body = candidate.email_body
        assert email_body is not None # to satisfy pylance

        # get relevant info for po_adjust, eg.:
        order_number_match = re.search(r"order_number:\s*(.+)", email_body)
        order_number = order_number_match.group(1) if order_number_match else None

        confirmed_qty_match = re.search(r"confirmed_qty:\s*(.+)", email_body)
        confirmed_qty = confirmed_qty_match.group(1) if confirmed_qty_match else None

        confirmed_qty = str(confirmed_qty)

        error_message = ""
        if not confirmed_qty.isnumeric() or int(confirmed_qty) < 0:
            error_message = f"invalid confirmed_qty={confirmed_qty}. "

        if error_message:
            return JobResult(is_success=False, error_message=error_message.strip())

        rpatool_payload = {
            "order_number": order_number,
            "confirmed_qty": confirmed_qty,
        }

        return JobResult(is_success=True, rpatool_payload=rpatool_payload)
    

    def verify_result(self, handover_file: HandoverFile) -> JobResult:
        # placeholder for implementation
        return JobResult(is_success=True)



def build_custom_shared_mail_handlers(logger) -> dict:
    return {
        "po_adjust": PoAdjustHandler(logger),
    }
