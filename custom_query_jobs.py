# place in main.py directory

from main import JobCandidate, JobResult, QueryWorkItem, JobName, HandoverFile, PostHandoverCrash

   
class OrderAdjustHandler:
    '''Example of a query driven automation for demo purpose'''

    job_name: JobName = "order_adjust"

    def __init__(self, logger, audit_repo, erp_backend) -> None:
        self.logger = logger
        self.audit_repo = audit_repo
        self.erp_backend = erp_backend


    def find_next_work_item(self) -> QueryWorkItem | None:
        rows = self.erp_backend.order_adjust_selection_rows()

        for row_raw in rows:
            candidate = self.erp_backend.build_candidate_from_row(row_raw)
            
            # Avoid reprocessing the same source_ref multiple times on the same day.
            if self.audit_repo.has_been_processed_today(candidate.source_ref):
                continue

            precheck_result = self.precheck_and_build_payload(candidate)

            if not precheck_result.is_success:
                self.logger.system(
                    f"query candidate rejected source_ref={candidate.source_ref}: {precheck_result.error_message}"
                )
                continue

            
            return QueryWorkItem(
                candidate=candidate,
                rpatool_payload=precheck_result.rpatool_payload or {},
            )

        return None

   
    def precheck_and_build_payload(self, candidate: JobCandidate) -> JobResult:
        source_ref = candidate.source_ref
        order_qty = candidate.parsed_source_data.get("order_qty")
        material_available = candidate.parsed_source_data.get("material_available")

        if order_qty == material_available:
            return JobResult(is_success=False, error_message="no mismatch left to fix")

        rpatool_payload = {
            "source_ref": str(source_ref),
            "target_order_qty": material_available,
        }

        return JobResult(is_success=True, rpatool_payload=rpatool_payload)
    

    def verify_result(self, handover_file: HandoverFile) -> JobResult:
        '''
        verify_result() must return:
        * success, or
        * failure with error_code VERIFICATION_MISMATCH or VERIFICATION_TIMEOUT
        all other outcomes are treated as programming/system fault. (implement eg. VERIFICATION_TIMEOUT if needed)
        '''
    
        job_id = handover_file.job_id

        rpatool_payload = handover_file.rpatool_payload
        if not rpatool_payload:
            raise PostHandoverCrash(message="missing rpatool_payload", job_id=job_id, handover_file=handover_file, cause=None)
        
        # get the order number/id and the target qty
        source_ref = rpatool_payload.get("source_ref")
        target_order_qty = rpatool_payload.get("target_order_qty")

        # get actual qty from ERP
        order_qty_erp = self.erp_backend.get_order_qty(source_ref)
        if order_qty_erp is None:
            return JobResult(
                is_success=False,
                error_code="VERIFICATION_TIMEOUT",
                error_message="ERP unreachable"
                )

        # compare them
        if order_qty_erp != target_order_qty:
            error_message= f"ERP shows mismatch. {source_ref} should be {target_order_qty}, is {order_qty_erp}"
            self.logger.system(error_message, job_id)
            return JobResult(
                is_success=False,
                error_code="VERIFICATION_MISMATCH",
                error_message=error_message
                )

        self.logger.system(f"OK. Should be: {target_order_qty}, is: {order_qty_erp}", job_id)
        return JobResult(
            is_success=True
            )


def build_custom_query_handlers(logger, audit_repo, erp_backend) -> dict:
    return {
        "order_adjust": OrderAdjustHandler(logger, audit_repo, erp_backend),
    }