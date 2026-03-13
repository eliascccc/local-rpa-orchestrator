#policy:
# Efter safestop/omstart av RPA/python är det alltid ett nytt kallt startläge.
# I produktion körs RPA på en windows-laptop utan admin-rättigheter, men dev sker i ubuntu, så koden behöver funka för båda. Python ver 3.14
# att döda processer är ok, jag kommer spara en lista över OK-processnamn taget när automationen är i full gång, och döda (matchat på namn) alla andra processer för att nollställa hela datorn efter varje jobb. Det är en dedikerat RPA-dator som inte ska ha massa pop-ups osv.
# normal system_state får endast ändras via write(). fatal system_state intern nödstoppstatus i Python
# email FÅR svälta schemalagda jobb, alltid prio på email.
# koden körs på en dedikerad RPA dator utan andra uppgifter. 
# max ett email per användare med "lifesign-notice" ska skickas per dag
# ett emailsvar ska skickas med antingen job done eller job failed till användare i friends.xls
# no resume policy: unfinished, paused or crashed jobs should not resume
# if python is in safestop, an operator may manually create reboot.flag, python then exits and creates reboot_ready.flag, after which RPA can restart main.py (this enables remote bug-fix and restart)
# this is the most simple and cheap set-up where you, as a team-member, will request an extra device (= no additional license cost for OS, Office etc.) and make it a dedicated RPA-machine.  
#Audit-status i SQLite ska namnen beskriva jobbets livscykel: RECEIVED, REJECTED (error by user, eg no access or invalid request), QUEUED (waiting for RPA), job_running, VERIFYING (double check with query), DONE, FAILED (error by robot, eg verification failed or crash)


# info about RPA:
# 1. On operator start of the external RPA-software it creates handover.txt with system_state "idle" and try removes 'stop.flag'
# 2. It runs this python script below async and then enters a while-true loop
# 3. within a TRY, the loop reads handover.txt and, and if read "queued", changes it to "job_running"
# 4. during 'job_running' it performs the automation in ERP, and when done changes handover.txt to "job_verifying"
# 5. any errors are catched en Except, that changes handover.txt to 'safestop'
# 6. all other states than 'queued' are ignored
# 7. On operator stop of the external RPA, it has a FINALLY-clause to create 'stop.flag'  



#ctrl + K + 2 för att collapsa alla metoder
# RPA automation framework with email-triggered job processing. 1 job at a time. beginner friendly. external RPA for screen-clicks.
import tkinter as tk
import time, random, threading, traceback, os, tempfile, sys, platform, subprocess, signal, atexit, sqlite3, datetime
from openpyxl import load_workbook #type:ignore
from typing import Never
from typing import Literal



job_state = Literal[
    "RECEIVED",        # email is recieved
    "QUEUED",          # job accepted and waiting for external robot
    "RUNNING",         # external robot executing
    "VERIFYING",       # verifying result
    "DONE",            # success
    "REJECTED",        # rejected before execution (user issue)
    "FAILED",          # robot/system error
]

system_state = Literal[
    "idle",
    "job_queued", 
    "job_running", 
    "job_verifying",
    "safestop",
]


class RecordingService:
    def __init__(self, append_system_log, ui, in_dev_mode) -> None:
        self.append_system_log = append_system_log
        self.ui = ui
        self.recording_process = None

        self.in_dev_mode = in_dev_mode


    def start(self, job_id) -> None:
        #written by AI
        

        if self.in_dev_mode:
            recording_process = None #remove in prod



        else: #remove in prod
            os.makedirs("recordings", exist_ok=True)
            filename = f"recordings/{job_id}.mkv"

            drawtext = (
                f"drawtext=text='job_id  {job_id}':"
                "x=200:y=20:"
                "fontsize=32:"
                "fontcolor=lightyellow:"
                "box=1:"
                "boxcolor=black@0.5"
            )

            if platform.system() == "Windows":
                capture = ["-f", "gdigrab", "-i", "desktop"]
                ffmpeg = "./ffmpeg.exe"
                recording_process = subprocess.Popen(
                    [ffmpeg, "-y", *capture, "-framerate", "15", "-vf", drawtext,
                    "-vcodec", "libx264", "-preset", "ultrafast", filename],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
                )
            else:
                capture = ["-video_size", "1920x1080", "-f", "x11grab", "-i", ":0.0"]
                ffmpeg = "ffmpeg"
                recording_process = subprocess.Popen(
                    [ffmpeg, "-y", *capture, "-framerate", "15", "-vf", drawtext,
                    "-vcodec", "libx264", "-preset", "ultrafast", filename],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True
                )
            time.sleep(0.2) #adding dummy time to start the recording
        
        self.append_system_log("recording started", job_id) if not self.in_dev_mode else self.append_system_log("DEV-mode: recording NOT STARTED", job_id)
        self.ui.root.after(0, self.ui.show_recording_overlay)
        self.recording_process = recording_process  
  

    def stop(self, job_id=None) -> None:
        #written by AI
        try:
            self.append_system_log(" ", job_id)
        except Exception: pass

        recording_process = self.recording_process
        self.recording_process = None

        try:
            if recording_process is not None:
                if platform.system() == "Windows":
                    try:
                        recording_process.send_signal(getattr(signal, "CTRL_BREAK_EVENT", signal.SIGTERM))
                    except Exception:
                        recording_process.terminate()

                    try:
                        recording_process.wait(timeout=8)
                    except subprocess.TimeoutExpired:
                        subprocess.run(
                            ["taskkill", "/IM", "ffmpeg.exe", "/T", "/F"],
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            check=False,
                        )

                else:
                    try:
                        os.killpg(recording_process.pid, signal.SIGINT)
                    except Exception:
                        recording_process.terminate()

                    try:
                        recording_process.wait(timeout=8)
                    except subprocess.TimeoutExpired:
                        subprocess.run(
                            ["killall", "-q", "-KILL", "ffmpeg"],
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            check=False,
                        )
            else:
                # fallback if proc-object tappats bort
                if platform.system() == "Windows":
                    subprocess.run(
                        ["taskkill", "/IM", "ffmpeg.exe", "/T", "/F"],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        check=False,
                    )
                else:
                    subprocess.run(
                        ["killall", "-q", "-KILL", "ffmpeg"],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        check=False,
                    )

        except Exception as err:
            print("WARN from stop():", err)

        finally:
            try: self.ui.root.after(0, self.ui.hide_recording_overlay)
            except Exception: pass


    def upload_recording_with_retry(self,local_file="jobid.mkv", remote_path="dummy", max_attempts=3):
        import shutil
        
        for attempt in range(max_attempts):
            try:
                # Create parent dir
                #remote_path.parent.mkdir(parents=True, exist_ok=True)
                
                # Copy file
               # shutil.copy2(local_file, remote_path)
                
                print(f"✓ Upload successful: {remote_path}")
                return True
            
            except Exception as e:
                wait_time = (attempt + 1) ** 2  # 1, 4, 9 seconds
                print(f"Attempt {attempt+1}/{max_attempts} failed: {e}")
                print(f"Retrying in {wait_time}s...")
                time.sleep(wait_time)
        
        return False


class Job1Handler:
    pass


class EmailJobHandler:
    def __init__(self, append_system_log, append_ui_log, friends_repo, update_ui_status, generate_job_id, is_within_operating_hours, network_monitor, recording_service, enter_safestop, handover_repo, audit_repo, in_dev_mode) -> None:
        
        self.in_dev_mode = in_dev_mode
        
        self.append_system_log = append_system_log
        self.append_ui_log = append_ui_log
        self.friends_repo = friends_repo
        self.update_ui_status = update_ui_status
        self.generate_job_id = generate_job_id
        self.is_within_operating_hours = is_within_operating_hours
        self.network_monitor = network_monitor
        self.recording_service = recording_service
        self.enter_safestop = enter_safestop
        self.handover_repo = handover_repo
        self.audit_repo = audit_repo
        
        self.fake_emails =["alice@example.com", "dummy,", "bob@test.com"]
        #self.fake_emails =[]
        self.in_dev_mode_use_fake_emails = True
        self.in_dev_mode_use_fake_emails_once = False #dont change this

       
    def process_inbox(self) -> bool:
        job_id=None
        received_notice_sent = False
        final_reply_sent = False

        emails = self.fetch_one_email_from_inbox()  #or many?
        for email_obj in emails:
            emails.remove(email_obj)

            if self.in_dev_mode_use_fake_emails_once: continue # remove in PROD

            email_address = email_obj # email_obj.sender
            email_subject = email_obj+"_subjekt" #email_obj.subject
            email_body = email_obj+"_body" #email_obj.body
            email_id = email_obj+"_id" #email_obj.id

            #sender, subj, body, email_id = _extract_email_fields()

            self.append_ui_log(f"email from {email_address}")
            self.append_system_log(f"new email from {email_address}")

            if not self.friends_repo.is_allowed_sender(email_address):
                self.delete_email(email_id)             #no reply
                self.append_system_log(f"email from {email_address} deleted (not in friends.xlsx)")
                self.append_ui_log("--> deleted (not in friends.xlsx)\n") 
                continue 


            #now we are busy
            sender_notified = False
            #self.python_is_busy = True

            self.update_ui_status("working")
            if self.in_dev_mode: time.sleep(2)

            try:
                job_id = self.generate_job_id()
                self.audit_repo.update_job_audit(job_id=job_id, job_status="RECEIVED", email_address=email_address, email_subject=email_subject, job_start_date=datetime.datetime.now().strftime("%Y-%m-%d"), job_start_time=datetime.datetime.now().strftime("%H:%M:%S"), insert_db_row=True)
   
                if not self.is_within_operating_hours():
                    self.reply_and_delete(email_id, job_id, message="Fail! email received outside working hours 07-20. Your email was deleted")
                    sender_notified = True
                    self.append_ui_log("--> delete & reply 'outside working hours'\n")
                    self.audit_repo.update_job_audit(job_id=job_id, job_status="REJECTED", error_code="OUTSIDE_WORKING_HOURS", error_explanation="email received outside working hours 07-20")
                    continue
            
                job_type = self.identify_email_job_type(email_id, job_id=job_id)

                if job_type == "unknown":
                    self.reply_and_delete(email_id, job_id, "Fail! Could not identify a job type from this email. Check spelling of keywords and/or attached files and send again.")
                    sender_notified = True
                    self.append_ui_log(f"--> delete & reply 'could not identify job type' \n")
                    self.audit_repo.update_job_audit(job_id=job_id, job_status="REJECTED", error_code="UNKNOWN_JOB", error_explanation=f"Unable to identify job type")
                    continue
                
                if not self.friends_repo.has_job_access(email_address=email_address, job_type=job_type):
                    self.reply_and_delete(email_id, job_id, f"FAIL! No access to {job_type}. Check with administrator for access.")
                    sender_notified = True
                    self.append_ui_log(f"--> delete & reply 'no access to {job_type} \n")
                    self.audit_repo.update_job_audit(job_id=job_id, job_status="REJECTED", error_code="NO_ACCESS", error_explanation=f"Sender has no access to {job_type}")
                    continue
                
                if not self.network_monitor.is_allowed_sender():
                    self.reply_and_delete(email_id, job_id, f"FAIL! No network. Try again later.")
                    sender_notified = True
                    self.append_ui_log(f"--> delete & reply 'no network' \n")
                    self.audit_repo.update_job_audit(job_id=job_id, job_status="REJECTED", error_code="NO_NETWORK", error_explanation=f"No network at the moment")
                    continue


                # --- special cases ---    (no handover)

                if job_type == "ping":
                    #playsound(ping.wav)
                    self.reply_and_delete(email_id, job_id, "PONG (robot online).")              
                    sender_notified = True
                    self.append_ui_log(f"--> delete & reply 'PONG' \n")
                    self.audit_repo.update_job_audit(job_id=job_id, job_status="DONE", job_finish_time=datetime.datetime.now().strftime("%H:%M:%S"))
                    continue 
                
                
                # --- standard pipeline ---   (with handover)

                if job_type == "job1":
                    is_valid, payload_or_error = self.precheck_job1(email_id)
                    if not is_valid:
                        error = payload_or_error
                        self.reply_and_delete(email_id, job_id, error)
                        sender_notified = True
                        self.append_ui_log(f"--> delete & reply 're-check input for {job_type}' \n")
                        self.audit_repo.update_job_audit(job_id=job_id, job_status="REJECTED", error_code="UNSPECIFIED", error_explanation=error)
                        del is_valid, error
                        continue
                
                elif job_type == "job2":
                    is_valid, payload_or_error = self.precheck_job2(email_id)
                    if not is_valid:
                        error = payload_or_error
                        self.reply_and_delete(email_id, job_id, error)
                        sender_notified = True                 
                        self.append_ui_log(f"--> delete & reply 're-check input for {job_type}' \n")
                        self.audit_repo.update_job_audit(job_id=job_id, job_status="REJECTED", error_code="UNSPECIFIED", error_explanation=error)
                        del error, is_valid
                        continue


                # --- mail accepted, now prepare for handover ---

                self.send_received_notice_if_first_today(email_address=email_address, job_id=job_id)
                sender_notified = True

                self.move_to_processing_folder(email_id)
                payload = payload_or_error  # required for standard pipeline
                
                try: self.recording_service.start(job_id)
                except Exception: raise RuntimeError("unable to start videorecording")
                
                self.audit_repo.update_job_audit(job_id=job_id, job_status="QUEUED")
                handover_data = {"system_state": "job_queued", "job_id": job_id, "job_type": job_type, "email_id": email_id, "created_at": time.time(),**payload}

                try:
                    self.handover_repo.write_file(handover_data)
                except Exception as err:
                    self.reply_and_delete(email_id, job_id, "FAIL! System error, your request is valid but could not start. Robot will stop (out-of-service) and your email was deleted. An automated email was sent to robot admin.")
                    sender_notified = True
                    self.append_ui_log(f"--> delete & reply 'system error' \n")
                    self.audit_repo.update_job_audit(job_id=job_id, job_status="FAILED", error_code="SYSTEM_ERROR", error_explanation=err)
                    self.enter_safestop(reason=err, job_id=job_id)
                    
                del handover_data, payload, payload_or_error
                
                if not sender_notified: raise ValueError("!!!!!!!!!!!!!!!!!!!!!!!!! add code to notify user")
                
                self.append_system_log(f"return True (RPA handover needed)", job_id)
                return True
            
            except Exception as err:
                try:
                    if not sender_notified:
                        self.reply_and_delete(email_id, job_id, "Unknown system error, the robot will do a full stop (out-of-service) and your email was deleted.")
                        sender_notified = True
                        self.append_ui_log(f"--> delete & reply 'unknown system error' \n")
                        self.audit_repo.update_job_audit(job_id=job_id, job_status="FAILED", error_code="SYSTEM_ERROR", error_explanation=err)
                except Exception: pass

                raise #notify sender when error and re-raise error
                    
        
        
        if self.in_dev_mode: self.in_dev_mode_use_fake_emails_once = True #remove in prod
        #self.append_system_log(f"return False (no unhandled emails in inbox)")
        return False


    def fetch_one_email_from_inbox(self):
        if self.in_dev_mode:
            return self.fake_emails
        else:
            return []


    def move_to_processing_folder(self,email):
        #move from inbox to "processing"
        pass


    def send_received_notice_if_first_today(self, job_id, email_address) -> None:
        

        jobs_today = self.audit_repo.count_todays_jobs_by_sender(job_id, email_address)

        if jobs_today != 0:
            return
          
         #under const.
        
        #rubrik RECEIVED re:
        ## This is an automated reply:
        # The robot is online(green) and your email is received.
        # Only one "RECEIVED"-email is sent per day to prevent spamming.
        #  A new email with the result (DONE/FAIL) will be sent when job is completed.
        pass
    

    def delete_email(self, email_id):
         #under construction
         pass
     

    def reply_and_delete(self, email_id, job_id, message):
        #under construction
        return

        #update also sqlite?

        self.append_system_log(f"arg: message='{message[:50]+'...' if len(message)>50 else message}'", job_id)
        
            # some "remove email" -action


    def identify_email_job_type(self,email_id, job_id) -> str:
        #add logic to identify job type
        job_type = "ping"
        job_type = "unknown"
        job_type = "job1"
        
        self.append_system_log(f"job_type is {job_type}", job_id)

        return(job_type)


    def precheck_job1(self, email_id) -> tuple[bool, dict]:

        payload_or_error = {"sku": 111, "old_material": 222}
        return True, payload_or_error


    def precheck_job2(self, email_id):
        return False,{}
    

class ScheduledJobHandler:
    def __init__(self, append_system_log, append_ui_log, update_ui_status, generate_job_id, is_within_operating_hours, network_monitor, recording_service, enter_safestop, handover_repo, audit_repo, in_dev_mode) -> None:
        self.append_system_log = append_system_log
        self.append_ui_log = append_ui_log
        self.update_ui_status = update_ui_status
        self.generate_job_id = generate_job_id
        self.is_within_operating_hours = is_within_operating_hours
        self.network_monitor = network_monitor
        self.recording_service = recording_service
        self.enter_safestop = enter_safestop
        self.handover_repo = handover_repo
        self.audit_repo = audit_repo
        self.count_todays_jobs_by_sender = audit_repo.count_todays_jobs_by_sender
        self.next_job3_check_time = 0
        self.next_job4_check_time = 0

        self.in_dev_mode = in_dev_mode


    def process_scheduled_jobs(self) -> bool:

        
        if not self.network_monitor.is_allowed_sender():
            return False
        
        now = time.time()
        
        #dispach
        if now > self.next_job3_check_time:
            found = self.process_scheduled_job3()
            if found:
                return True #yes handover needed        
            self.next_job3_check_time = now + 3600 #1h
            if self.in_dev_mode: self.next_job3_check_time = now + 10
        
        if now > self.next_job4_check_time:
            found = self.process_scheduled_job4()
            if found:
                return True #yes handover needed
            self.next_job4_check_time = now + 3600 #1h
            if self.in_dev_mode: self.next_job4_check_time = now + 10

        #self.append_system_log(f"return False (no scheduled jobs found)")
        return False #no handover needed
    

    def unused_simulate_a_new_job3(self):
        with open("job3.flag", "w", encoding="utf-8") as f:
            f.write("simuation of workflow: job1")


    def process_scheduled_job3(self) -> bool:
        self.append_system_log(f"check")
        
        if self.in_dev_mode:
            return False
        
        if random.randint(0, 1) == 1: #simulation
            return False #no jobs found

        
        #now we busy
        self.update_ui_status(status="working")
        
        job_id = self.generate_job_id()
        job_type="job3"
        self.append_system_log("new job found, job_id created", job_id)
        self.append_ui_log("job found!")
        self.audit_repo.update_job_audit(job_id=job_id, job_status="QUEUED", job_type=job_type, job_start_date=datetime.datetime.now().strftime("%Y-%m-%d"), job_start_time=datetime.datetime.now().strftime("%H:%M:%S"), insert_db_row=True)


        handover_data = {"system_state": "job_queued", "job_id":job_id, "job_type": job_type}
        self.handover_repo.write_file(handover_data)
        return True
        

    def process_scheduled_job4(self) -> bool:
        #placeholder for job4 logic
        self.append_system_log("check job4")
        return False


class JobVerifier:
    """this class handels verifaication of both email jobs and scheduled jobs"""
    def __init__(self, append_system_log, ui, audit_repo, email_job_handler, recording_service, handover_repo) -> None:
        self.append_system_log=append_system_log
        self.ui=ui
        self.audit_repo=audit_repo
        self.email_job_handler = email_job_handler
        self.recording_service = recording_service
        self.handover_repo = handover_repo

    
    def process_verification(self, handover_data):
        #a dubbel-check that the intended entered handover_data in ERP via RPA is indeed entered according to a query 
        job_id = handover_data.get("job_id")
        job_type = handover_data.get("job_type")

        self.append_system_log(f"fetched: {handover_data}", job_id)
        self.audit_repo.update_job_audit(job_id=job_id, job_finish_time=datetime.datetime.now().strftime("%H:%M:%S"), job_status="VERIFYING") 



        if job_type == "job1":
            self.verify_job1()
            
        elif job_type == "job2":
            self.verify_job2()

        elif job_type == "job3":
            #add logic
            pass

        elif job_type == "job4":
            #add logic
            pass


        time.sleep(3) #simulate job_verifying
        self.audit_repo.update_job_audit(job_id=job_id, job_finish_time=datetime.datetime.now().strftime("%H:%M:%S"), job_status="DONE") 
        self.append_system_log("success", job_id)
        self.complete_verification(job_id)


    def verify_job1(self):
        """Query ERP to confirm job1 was entered correctly"""
        pass


    def verify_job2(self):
        pass


    def complete_verification(self, job_id) -> None:
        self.recording_service.stop(job_id)

        count = self.audit_repo.count_completed_jobs_today()
        self.ui.root.after(0, lambda: self.ui.set_jobs_done_today(count))

        self.handover_repo.write({"system_state": "idle"})
        self.append_system_log("state: job_verifying -> idle", job_id)
        time.sleep(0.1) # to allow update_status_display: online
  

class HandoverRepository:
   
    def __init__(self, append_system_log) -> None:
        self.append_system_log = append_system_log
        self.VALID_JOB_TYPES = ("job1", "job2", "job3", "job4")
        self.VALID_SYSTEM_STATES = ("idle", "job_queued", "job_running", "job_verifying", "safestop")

   
    def read_file(self) -> dict:
        last_err=None

        for attempt in range(7):
            try:
                handover_data = {}
                with open("handover.txt", "r", encoding="utf-8") as f:
                    for row in f:
                        row = row.strip()
                        if not row: continue
                        if "=" not in row: raise ValueError(f"Invalid row in handover: {row}")
                        key, value = row.split("=", 1)
                        handover_data[key.strip()] = value.strip()

                system_state = handover_data.get("system_state")               # validate state
                if system_state not in self.VALID_SYSTEM_STATES:
                    raise ValueError(f"Unknown state: {system_state}")
                
                job_type = handover_data.get("job_type")    #validate job_type
                if system_state =="job_verifying" and job_type not in self.VALID_JOB_TYPES:
                    raise ValueError(f"Unknown job_type for system_state job_verifying: {job_type}")

                return handover_data

            except Exception as err:
                last_err = err
                print(f"WARN: retry {attempt+1}/7 : {err}")
                time.sleep((attempt+1) ** 2)


        raise RuntimeError(f"handover.txt unreadable: {last_err}")
    
      
    def write_file(self, handover_data: dict) -> None:
        """ atomic write of handover.txt"""

        system_state = handover_data.get("system_state")               # validate state
        if system_state not in self.VALID_SYSTEM_STATES:
            raise ValueError(f"Unknown state: {system_state}")

        for attempt in range(7):
            temp_path = None
            try:
                dir_path = os.path.dirname(os.path.abspath("handover.txt"))
                fd, temp_path = tempfile.mkstemp(dir=dir_path)    # create temp file

                #atomic write
                with os.fdopen(fd, "w", encoding="utf-8") as tmp:
                    for key, value in handover_data.items():
                        if value is None: value = ""
                        tmp.write(f"{key}={value}\n")
                    tmp.flush()
                    os.fsync(tmp.fileno())

                os.replace(temp_path, "handover.txt") # replace original file
                try: self.append_system_log(f"written: {handover_data}", job_id=handover_data.get("job_id"))
                except Exception: pass
                return

            except Exception as err:
                last_err = err
                print(f"{attempt+1}st warning from write()")
                try: self.append_system_log(f"WARN: {attempt+1}/7 error", job_id=handover_data.get("job_id"))
                except Exception: pass
                time.sleep((attempt + 1) ** 2) # 1 4 9 ... 49sec

            finally: #remove temp-file if writing fails.
                if temp_path and os.path.exists(temp_path):
                    try: os.remove(temp_path)
                    except Exception: pass

        try: self.append_system_log(f"CRITICAL: cannot write handover.txt {last_err}", job_id=handover_data.get("job_id"))
        except Exception: pass
        raise RuntimeError("CRITICAL: cannot write handover.txt")
  

class FriendsRepository:
    def __init__(self, append_system_log) -> None:
        self.append_system_log = append_system_log
        self.friends_access = {}
        self.friends_file_mtime = None
    

    def read_access_file(self, filepath="friends.xlsx") -> dict:
        #code written by AI
        """
        Reads friends.xlsx and returns eg.:

        {
            "alice@example.com": {"ping"},
            "ex2@whatever.com": {"ping", "job1"}
        }

        Presumptions:
        A1 = email
        row 1 contains job_type
        'x' gives access
    
        """
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active

        rows = list(ws.iter_rows(values_only=True)) # type: ignore
        if len(rows) < 2:
            raise ValueError("friends.xlsx contains no users")

        header = rows[0]   # första raden

        access_map: dict[str, set[str]] = {}

        for row in rows[1:]:
            email_cell = row[0]

            if email_cell is None:
                continue

            email = str(email_cell).strip().lower()
            if not email:
                continue

            permissions = set()

            for col in range(1, len(header)):

                jobname = header[col]
                if jobname is None:
                    continue

                jobname = str(jobname).strip().lower()

                cell = row[col] if col < len(row) else None
                if cell is None:
                    continue

                if str(cell).strip().lower() == "x":
                    permissions.add(jobname)

            access_map[email] = permissions

        return access_map


    def refresh_friends_access(self, force_reload=False, filepath="friends.xlsx") -> bool:
        #code written by AI
        """
        Laddar om friends.xlsx om filen ändrats sedan sist.
        force_reload=True tvingar omladdning.
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"{filepath} not found")

        current_mtime = os.path.getmtime(filepath)

        if (not force_reload) and (self.friends_file_mtime == current_mtime):
            return False   # ingen ändring

        new_access = self.read_access_file(filepath)

        self.friends_access = new_access
        self.friends_file_mtime = current_mtime

        return True


    def is_allowed_sender(self, email_address: str) -> bool:
        email_address = email_address.strip().lower()
        result = email_address in self.friends_access
        return result


    def has_job_access(self, email_address: str, job_type: str) -> bool:
        email_address = email_address.strip().lower()
        job_type = job_type.strip().lower()
        result = job_type in self.friends_access.get(email_address, set())
        return result


class AuditRepository:
    def __init__(self, append_system_log) -> None:
        self.append_system_log = append_system_log
        
        # Connect
        self.db = sqlite3.connect("log.db")
        self.db.row_factory = sqlite3.Row

    def create_db_if_first_run(self) -> None:

        with sqlite3.connect("log.db") as conn:
            cur = conn.cursor()

            cur.execute("""
                CREATE TABLE IF NOT EXISTS audit_log
                         (job_id INTEGER PRIMARY KEY, email_address TEXT, email_subject TEXT, job_type TEXT, job_start_date TEXT, job_start_time TEXT, job_finish_time TEXT, job_status TEXT, error_code TEXT, error_explanation TEXT )
                        """)


    def update_job_audit(self, job_id, email_address=None, email_subject=None, job_type=None, job_start_date=None, job_start_time=None, job_finish_time=None, job_status=None, error_code=None,error_explanation=None, insert_db_row=False) -> None:
        # example use: self.audit_repo.update_job_audit(job_id=20260311124501, job_type="job1")

        if job_status not in ("RECEIVED", "REJECTED", "QUEUED", "RUNNING", "VERIFYING", "DONE", "FAILED", None):
            raise ValueError(f"update_job_audit(): unknown job_status={job_status}")

        all_fields = {
            "job_id": job_id,
            "email_address": email_address,
            "email_subject": email_subject,
            "job_type": job_type,
            "job_start_date": job_start_date,
            "job_start_time": job_start_time,
            "job_finish_time": job_finish_time,
            "job_status": job_status,
            "error_code": error_code,
            "error_explanation": error_explanation,
        }

        fields = {k: v for k, v in all_fields.items() if v is not None}
        self.append_system_log(f"appending: {fields}", job_id=job_id)

        with sqlite3.connect("log.db") as conn:
            cur = conn.cursor()

            if insert_db_row:
                columns = ", ".join(fields.keys())
                placeholders = ", ".join("?" for _ in fields)

                cur.execute(
                    f"INSERT INTO audit_log ({columns}) VALUES ({placeholders})",
                    tuple(fields.values())
                )

            else:
                fields.pop("job_id", None)

                if not fields:
                    return

                set_clause = ", ".join(f"{k}=?" for k in fields)

                cur.execute(
                    f"UPDATE audit_log SET {set_clause} WHERE job_id=?",
                    (*fields.values(), job_id)
                )

                if cur.rowcount == 0:
                    raise ValueError(f"update_job_audit(): no row in DB with job_id={job_id}")
    

    def count_completed_jobs_today(self) -> int:
        if not os.path.isfile("log.db"):
            return 0

        today = datetime.date.today().isoformat()

        with sqlite3.connect("log.db") as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT COUNT(*)
                FROM audit_log
                WHERE job_start_date = ?
                AND job_status = 'DONE'
            """, (today,))
            
            result = cur.fetchone()[0]
            return result


    def count_todays_jobs_by_sender(self, job_id, email_address) -> int:    
        #query for previuos jobs today from sender
        #self.append_system_log("running")

        today = datetime.datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect("log.db")
        cur = conn.cursor()

        cur.execute(
            """
            SELECT COUNT(*)
            FROM audit_log
            WHERE job_id != ? AND job_start_date = ? AND email_address = ?
            """,
            (job_id, today, email_address,)
        )

        jobs_today = cur.fetchone()[0]
        conn.close()

        return jobs_today


    #extra below unused so far.. keep or remove ?

    # Recent jobs
    def get_recent_jobs(self,limit=10):
        cur = self.db.cursor()
        cur.execute("""
            SELECT job_id, email_sender, job_type, job_status, job_finish_time
            FROM audit_log
            ORDER BY job_id DESC
            LIMIT ?
        """, (limit,))
        return cur.fetchall()

    # Jobs by sender
    def get_jobs_by_sender(self, email_sender):
        cur = self. db.cursor()
        cur.execute("""
            SELECT job_id, job_type, job_status, job_finish_time
            FROM audit_log
            WHERE email_sender = ?
            ORDER BY job_id DESC
        """, (email_sender,))
        return cur.fetchall()

    # Failed jobs
    def get_failed_jobs(self, days=7):
        cur = self.db.cursor()
        cur.execute("""
            SELECT job_id, email_sender, job_type, error_code, error_explanation
            FROM audit_log
            WHERE job_status = 'FAILED'
            AND job_start_date >= date('now', '-' || ? || ' days')
            ORDER BY job_id DESC
        """, (days,))
        return cur.fetchall()

    # Jobs today
    def get_jobs_today(self):
        cur = self.db.cursor()
        cur.execute("""
            SELECT COUNT(*) as count, job_status
            FROM audit_log
            WHERE job_start_date = date('now')
            GROUP BY job_status
        """)
        return cur.fetchall()

    # Display in UI
    def get_dashboard_stats(self):
        """Stats for dashboard display"""
        db = sqlite3.connect("log.db")
        db.row_factory = sqlite3.Row
        
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        
        # Today's summary
        cur = db.cursor()
        cur.execute("""
            SELECT 
                SUM(CASE WHEN job_status='DONE' THEN 1 ELSE 0 END) as done,
                SUM(CASE WHEN job_status='FAILED' THEN 1 ELSE 0 END) as failed,
                SUM(CASE WHEN job_status='REJECTED' THEN 1 ELSE 0 END) as rejected,
                SUM(CASE WHEN job_status='RPA_RUNNING' THEN 1 ELSE 0 END) as running
            FROM audit_log
            WHERE job_start_date = ?
        """, (today,))
        
        row = cur.fetchone()
        return {
            "completed": row['done'] or 0,
            "failed": row['failed'] or 0,
            "rejected": row['rejected'] or 0,
            "in_progress": row['running'] or 0,
        }


class NetworkMonitor:
    def __init__(self, append_system_log, in_dev_mode) -> None:
        self.append_system_log = append_system_log
        self.network_state = None
        self.next_network_check_time = 0
        self.NETWORK_TEST_PATH = r"/" #enter path to network drive here, e.g. "G:\"

        self.in_dev_mode = in_dev_mode


    def is_allowed_sender(self) -> bool:
        #this runs at highest every hour, or before new jobs

        now = time.time()

        if now < self.next_network_check_time:
            return True if (self.network_state is not False) else False

        try:
            os.listdir(self.NETWORK_TEST_PATH)
            online = True
            
        except Exception:
            online = False
            

        # logga / uppdatera UI bara vid förändring
        if online != self.network_state:
            self.network_state = online

            if online:
                self.append_system_log("network restored")
            else:
                self.append_system_log(f"WARN: network lost")

        # olika pollingintervall beroende på status
        if online:
            self.next_network_check_time = now + 3600   # 1 h
        else:
            self.next_network_check_time = now + 60     # 1 min
        
        return online



########## Below are the three main-classes ##########################

#Core automation logic with job processing pipeline
class RobotRuntime:

    def __init__(self, ui):

        self.in_dev_mode = True

        self.ui = ui
        self.handover_repo = HandoverRepository(append_system_log=self.append_system_log)  #nu äger Runtime en handover_repo (som får med append-metod)
        self.friends_repo = FriendsRepository(append_system_log=self.append_system_log)
        self.network_monitor = NetworkMonitor(append_system_log=self.append_system_log, in_dev_mode=self.in_dev_mode)
        self.recording_service = RecordingService(append_system_log=self.append_system_log, ui=self.ui, in_dev_mode=self.in_dev_mode) #nu har vi skapar EN instans av denna klass, och vi kan skicka denna instans både till email-jobb, verification, och scheduled-job
        self.audit_repo = AuditRepository(append_system_log=self.append_system_log)
        self.email_job_handler = EmailJobHandler(append_system_log=self.append_system_log, append_ui_log=self.append_ui_log, friends_repo=self.friends_repo, update_ui_status=self.update_ui_status, generate_job_id=self.generate_job_id, is_within_operating_hours=self.is_within_operating_hours, network_monitor=self.network_monitor, recording_service=self.recording_service, enter_safestop=self.enter_safestop, handover_repo=self.handover_repo, audit_repo=self.audit_repo, in_dev_mode=self.in_dev_mode)
        self.scheduled_job_handler = ScheduledJobHandler(append_system_log=self.append_system_log, append_ui_log=self.append_ui_log, update_ui_status=self.update_ui_status, generate_job_id=self.generate_job_id, is_within_operating_hours=self.is_within_operating_hours, network_monitor=self.network_monitor, recording_service=self.recording_service, enter_safestop=self.enter_safestop, handover_repo=self.handover_repo, audit_repo=self.audit_repo, in_dev_mode=self.in_dev_mode)
        self.job_verifier = JobVerifier(append_system_log=self.append_system_log, ui=self.ui, audit_repo=self.audit_repo, email_job_handler = self.email_job_handler, recording_service=self.recording_service, handover_repo=self.handover_repo)
        self._safestop_entered = False    

    
    def initialize_runtime(self):
        if self.in_dev_mode: self.handover_repo.write_file({"system_state":"idle"}) # when deployed RPA will do this before running main.py


        VERSION = 0.4
        self.append_system_log(f"RuntimeThread started, version={VERSION}")

        # cleanup
        for fn in ["stop.flag", "reboot.flag", "reboot_ready.flag"]:
            try: os.remove(fn)
            except Exception: pass

        atexit.register(self.recording_service.stop) #extra protection during normal python exit
        self.recording_service.stop() #stop any remaing recordings 
        self.network_monitor.is_allowed_sender()

        try: self.friends_repo.refresh_friends_access(force_reload=True)
        except Exception as err: self.enter_safestop(reason=err)

        try: self.audit_repo.create_db_if_first_run()
        except Exception as err: self.enter_safestop(reason=err)

        try:
            count = self.audit_repo.count_completed_jobs_today()
            self.ui.root.after(0, lambda: self.ui.set_jobs_done_today(count))
        except Exception as err: self.enter_safestop(reason=err)

        #self._init_audit_db()


    def update_ui_status(self, arg=None) -> None:
        try:
            handover_data = self.handover_repo.read_file()
            system_state = handover_data.get("system_state")
            print("poll handover.txt from refresh_ui...()")
        except Exception:
            system_state = None

        if self._safestop_entered or system_state == "safestop":
            ui_status = "safestop"

        elif arg == "working" or system_state in ("job_queued", "job_running", "job_verifying"):
            ui_status = "working"

        elif self.network_monitor.network_state is False:
            ui_status = "no network"

        elif not self.is_within_operating_hours():
            ui_status = "ooo"

        else:
            ui_status = "online"

        self.ui.root.after(0, lambda: self.ui.update_status_display(ui_status))


    def run(self) -> None:
        self.initialize_runtime()

        sleep_s = 1       
        watchdog_timeout = 600 #10 min

        prev_system_state = None
        rpa_stalled_since = None

        while True:
            try:
                handover_data = self.handover_repo.read_file()
                system_state = handover_data.get("system_state")
                
                #dispatch
                if system_state == "idle":
                    self.check_for_jobs()  #set system_state=job_queued if RPA proccessing needed
                    time.sleep(sleep_s)

                elif system_state == "job_queued":  #RPA-poll trigger
                    time.sleep(sleep_s)

                elif system_state == "job_running":  #RPA set system_state=job_running when job fetched
                    time.sleep(sleep_s)

                elif system_state == "job_verifying":       #RPA set system_state=job_verifying when job completed
                    self.job_verifier.process_verification(handover_data) 

                elif system_state == "safestop":  #only RPA can trigger 'safestop' this way 
                    self.email_job_handler.reply_and_delete(email_id=handover_data.get("email_id"), job_id=handover_data.get("job_id"), message="rpa crash")
                    self.enter_safestop(reason="RPA safestop", job_id=handover_data.get("job_id"))
                    

                #log all system_state transitions
                if system_state != prev_system_state:
                    self.append_system_log(f"state transition detected by CPU-poll: {prev_system_state} -> {system_state}")
                    print("state is", system_state)

                    #update DB and set hang_timer
                    if system_state == "job_queued":
                        rpa_stalled_since = time.time()  #time.time() is float for 'how many seconds have passed this epoch'
                    elif system_state == "job_running":
                        rpa_stalled_since = time.time()
                        self.audit_repo.update_job_audit(job_id=handover_data.get("job_id"), job_status="RUNNING")
                    else:
                        rpa_stalled_since = None

                #detect RPA hang (actually: if no system_state transition from RPA )
                if rpa_stalled_since and system_state in ("job_queued", "job_running") and time.time() - rpa_stalled_since > watchdog_timeout:
                    self.email_job_handler.reply_and_delete(email_id=handover_data.get("email_id"), job_id=handover_data.get("job_id"), message="FAILED. This is a timeout error, very bad error, and you _NEED_ to watch video")
                    rpa_stalled_since = None
                    self.enter_safestop(reason="RPA timeout - no progress for 10 min", job_id=handover_data.get("job_id"))
                    
                
                prev_system_state = system_state
                #self.python_is_busy = False
                self.update_ui_status()
                #print(".", end="", flush=True)
                 
            except Exception:
                reason = traceback.format_exc()                
                self.enter_safestop(reason=reason)             


    def append_ui_log(self, text:str) -> None:
        #wrapper
        self.ui.root.after(0, lambda: self.ui.append_log_line(text))


    def append_system_log(self, event_text: str, job_id=None):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") #+ "."+str(datetime.datetime.now().microsecond)[:-4]
        try: caller_name = sys._getframe(1).f_code.co_name
        except Exception: caller_name ="caller_name error"
        job_part = f" | JOB {job_id}" if job_id else ""
        log_line = f"{timestamp} | PY{job_part} | {caller_name}() {event_text}"
        
        try:
            self.append_ui_log(log_line) #system.log is more important
        except Exception:
            pass
        
        last_err = None
        for i in range(7):
            try:
                with open("system.log", "a", encoding="utf-8") as f:
                    f.write(log_line + "\n")
                    f.flush()
                    os.fsync(f.fileno())
                return 

            except Exception as err:
                last_err = err
                print(f"WARN: retry {i+1}/7 from append_system_log():", err)
                if not self.in_dev_mode: time.sleep(i+1) #fail fast in dev

        raise RuntimeError(f"append_system_log() failed after 7 attempts: {last_err}")

 
    def enter_safestop(self, reason, job_id=None) -> Never | None:
        #critical errors and crashes end up here

        print("ROBOTRUNTIME CRASHED:\n", reason)

        
        if self._safestop_entered: return #re-entrancy protection
        self._safestop_entered = True 

        try: self.append_system_log(f"CRASH: {reason}", job_id)
        except Exception: pass

        try: self.send_admin_alert(reason)
        except Exception: pass

        try: self.append_ui_log("Error-email sent to admin. All automations halted")
        except Exception: pass

        try: self.recording_service.stop()
        except Exception: pass

        try: self.ui.root.after(0, lambda: self.ui.update_status_display("safestop")) 
        except Exception:
            print("unable to set ui_status text to 'safestop', shutting down...")
            try:
                self.ui.root.after(0, lambda: self.ui.shutdown())
            except:
                print("unable to soft-shutdown. Forcing exit")
                os._exit(1)
            time.sleep(3)
            os._exit(1)  #kill if still alive after 3 sec soft-shutdown 

        self.wait_for_reboot_after_safestop(job_id)
    

    def wait_for_reboot_after_safestop(self, job_id) -> Never:
        #experimental
        try: self.append_system_log("running", job_id)
        except Exception: pass

        while True:
            time.sleep(1)
            try:
                if os.path.isfile("reboot.flag"):
                    os.remove("reboot.flag")
                    open("reboot_ready.flag", "w").close()
                    os._exit(1)
            except Exception: pass


    def send_admin_alert(self, reason):
        # add logic to email admin
        pass


    def check_for_jobs(self) -> None:
        
        if self.friends_repo.refresh_friends_access(): self.append_system_log(f"friends.xlsx reloaded")

        
        #return stops further checks
        did_RPA_handover = self.email_job_handler.process_inbox()
        if did_RPA_handover:
            return

        did_RPA_handover = self.scheduled_job_handler.process_scheduled_jobs()
        if did_RPA_handover:
            return
        
        #placeholder other tasks
        
 
    def generate_job_id(self) -> int:
        """ makes the unique value assosiated with this job """
        time.sleep(1) # shady dublicate-value prevention

        job_id = int(datetime.datetime.now().strftime("%Y%m%d%H%M%S"))
        return job_id

    
    def is_within_operating_hours(self) -> bool:
        #return False
        now = datetime.datetime.now().time()
        return datetime.time(5,0) <= now <= datetime.time(23,0)
    

    def poll_for_stop_flag(self):
        print("poll_for_stop_flag() alive")
        while True:
            time.sleep(1)
            if os.path.isfile("stop.flag"):
                try: os.remove("stop.flag")
                except Exception: pass
                print("stop.flag found, requesting showdown")
                
                try: self.ui.root.after(0, lambda: self.ui.shutdown()) #request soft-exit from g if possible
                except Exception: os._exit(1)
                
                time.sleep(3)
                os._exit(1)  #kill if still alive after 3 sec soft-exit 


class DashboardUI:
    # Tkinter dashboard for monitoring
    def __init__(self):
        bg_color ="#000000" #or "#111827"
        text_color = "#F5F5F5"

        self._build_root(bg_color)
        self._build_header(bg_color, text_color)
        self._build_body(bg_color, text_color)
        self._build_footer(bg_color, text_color)
        
        #self.debug_grid(self.root)
    
    def attach_runtime(self, robot_runtime) -> None:
        self.robot_runtime = robot_runtime

    def run(self) -> None:
        self.root.mainloop()


    def _build_root(self,bg_color):
        self.root = tk.Tk()
        self.root.geometry('1800x1000+0+0')
        #self.root.geometry('1800x200+0+0')
        #self.root.attributes("-fullscreen", True)
        self.root.resizable(False, False)

        self.root.configure(bg=bg_color, padx=50)
        self._closing = False
        self.root.protocol("WM_DELETE_WINDOW", self.shutdown)

        self.root.title('RPA dashboard')
        self._create_recording_overlay()

                # --- Layout: root uses grid ---
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)


    def _build_header(self, bg_color, text_color):
        self.header = tk.Frame(self.root, bg=bg_color)
        
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.grid_columnconfigure(2, weight=1)  
        self.header.grid_rowconfigure(0, weight=1)  

               # --- Header content ---
        self.rpa_text_label = tk.Label(self.header, text="RPA:", fg=text_color, bg=bg_color, font=("Arial", 100, "bold"))  #snyggare: "Segoe UI"
        self.rpa_text_label.grid(row=0, column=0, padx=16, pady=16, sticky="w")
        self.rpa_status_label = tk.Label(self.header, text="", fg="red", bg=bg_color, font=("Arial", 100, "bold"))
        self.rpa_status_label.grid(row=0, column=1, padx=16, pady=16, sticky="w")
        self.status_dot = tk.Label(self.header, text="", fg="#22C55E", bg=bg_color, font=("Arial", 50, "bold"))
        self.status_dot.grid(row=0, column=2, sticky="w")


        # --- Jobs done today (counter + label in same) ---
        self.jobs_counter_frame = tk.Frame(self.header, bg=bg_color)
        self.jobs_counter_frame.grid(row=0, column=3, sticky="ne", padx=40, pady=30)
        self.jobs_counter_frame.grid_rowconfigure(0, weight=1)
        self.jobs_counter_frame.grid_columnconfigure(0, weight=1)


        # --- NORMAL VIEW (jobs done today) ---
        
        self.jobs_normal_view = tk.Frame(self.jobs_counter_frame, bg=bg_color)
        self.jobs_normal_view.grid(row=0, column=0, sticky="nsew")
        self.jobs_normal_view.grid_columnconfigure(0, weight=1)

        self.jobs_done_label = tk.Label(    self.jobs_normal_view,    text="0",    fg=text_color,    bg=bg_color,    font=("Segoe UI", 140, "bold"),       anchor="e",        justify="right")
        self.jobs_done_label.grid(row=0, column=0, sticky="e")

        self.jobs_counter_text = tk.Label(            self.jobs_normal_view,            text="jobs done today",            fg="#A0A0A0",            bg=bg_color,            font=("Arial", 14, "bold"),            anchor="e"        )
        self.jobs_counter_text.grid(row=1, column=0, sticky="e", pady=(0, 6))

        # --- SAFESTOP VIEW (stort X) ---
        self.jobs_error_view = tk.Frame(self.jobs_counter_frame, bg=bg_color)
        self.jobs_error_view.grid(row=0, column=0, sticky="nsew")

        self.safestop_x_label = tk.Label(            self.jobs_error_view,                        text="X",            bg="#DC2626",            fg="#FFFFFF",            font=("Segoe UI", 140, "bold")        ) #text="✖",
        self.safestop_x_label.pack(expand=True)


        # show normal view at startup
        self.jobs_normal_view.tkraise()

        #online-status animation
        self._online_animation_after_id = None
        self._online_pulse_index = 0

        #"working..."-status animation
        self._working_animation_after_id = None
        self._working_dots = 0


    def _build_body(self,bg_color, text_color):
        self.body = tk.Frame(self.root, bg=bg_color)        
        self.body.grid(row=1, column=0, sticky="nsew")
        self.body.grid_rowconfigure(0, weight=1)
        self.body.grid_columnconfigure(0, weight=1)

                        # --- Body content ---
        log_and_scroll_container = tk.Frame(self.body, bg=bg_color)
        log_and_scroll_container.grid(row=0, column=0, sticky="nsew")
        log_and_scroll_container.grid_rowconfigure(0, weight=1)
        log_and_scroll_container.grid_columnconfigure(0, weight=1)

        #the right-hand side scrollbar
        scrollbar = tk.Scrollbar(log_and_scroll_container, width=23, troughcolor="#0F172A", bg="#1E293B", activebackground="#475569", bd=0, highlightthickness=0, relief="flat")
        scrollbar.grid(row=0, column=1, sticky="ns")

        #the 'console'
        self.log_text = tk.Text(log_and_scroll_container, yscrollcommand=scrollbar.set, bg=bg_color, fg=text_color, insertbackground="black", font=("DejaVu Sans Mono", 10), wrap="none", state="disabled", bd=0,highlightthickness=0) #glow highlightbackground="#1F2937", highlightthickness=1   ## font=("DejaVu Sans Mono", 35)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.config(command=self.log_text.yview)


    def _build_footer(self,bg_color, text_color):
        self.footer = tk.Frame(self.root, bg=bg_color)        
        self.footer.grid(row=2, column=0, sticky="nsew")
        self.footer.grid_rowconfigure(0, weight=1)
        self.footer.grid_columnconfigure(0, weight=1)
        
                        # ---- Footer content ---
        self.last_activity_label = tk.Label(self.footer, text="last activity: 11:56", fg="#A0A0A0", bg=bg_color, font=("Arial", 14, "bold"), anchor="e")
        self.last_activity_label.grid(row=0, column=1, padx=8, pady=16)
        
        #remove this button?
        self.extended_log_button = tk.Button(self.footer, text="toggle extended log", bg="#2c3d2c", font=("Arial", 14, "bold"), command=self.do_something)
        self.extended_log_button.grid(row=0, column=4, padx=8, pady=16)


    def debug_grid(self,widget):
        #highlights all gris with red
        for child in widget.winfo_children():
            try:
                child.configure(highlightbackground="red", highlightthickness=1)
            except Exception:
                pass
            self.debug_grid(child)


    def update_status_display(self, status=None):
        #sets the status

        #stops any ongoing animations
        self._stop_online_animation()
        self._stop_working_animation()
        self.status_dot.config(text="")


        #changes text
        if status=="online":
            self.rpa_status_label.config(text="online", fg="#22C55E")
            self.jobs_normal_view.tkraise()
            self.status_dot.config(text="●")
            self._start_online_animation()
            
        elif status=="no network":
            self.rpa_status_label.config(text="no network", fg="red")
            self.jobs_normal_view.tkraise()
            
        elif status=="working":
            self.rpa_status_label.config(text="working...", fg="#FACC15")
            self.jobs_normal_view.tkraise()
            self._start_working_animation()

        elif status=="safestop":
            self.rpa_status_label.config(text="safestop", fg="red")
            self.jobs_error_view.tkraise()
            
        elif status=="ooo":
            self.rpa_status_label.config(text="out-of-office", fg="#FACC15")
            self.jobs_normal_view.tkraise()


    def set_jobs_done_today(self, n) -> None:
        self.jobs_done_label.config(text=str(n))


    def _create_recording_overlay(self) -> None:
        #written by AI
        self.recording_win = tk.Toplevel(self.root)
        self.recording_win.withdraw()                 # hidden at start
        self.recording_win.overrideredirect(True)    # no title/boarder
        self.recording_win.configure(bg="black")

        try: self.recording_win.attributes("-topmost", True)
        except Exception: pass

        width = 250
        height = 110
        x = self.root.winfo_screenwidth() - width - 30
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.recording_win.geometry(f"{width}x{height}+{x}+{y}")

        frame = tk.Frame(           self.recording_win,            bg="black",            highlightbackground="#444444",            highlightthickness=1,            bd=0        )
        frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(        frame,        width=44,        height=44,        bg="black",        highlightthickness=0,        bd=0        )
        canvas.place(x=18, y=33)
        canvas.create_oval(4, 4, 40, 40, fill="#DC2626", outline="#DC2626")

        label = tk.Label(            frame,            text="RECORDING",            fg="#FFFFFF",            bg="black",            font=("Arial", 20, "bold"),            anchor="w"        )
        label.place(x=75, y=33)

        
    def show_recording_overlay(self) -> None:
        #written by AI
        try:
            width = 250
            height = 110
            x = self.root.winfo_screenwidth() - width - 30
            y = (self.root.winfo_screenheight() // 2) - (height // 2)
            self.recording_win.geometry(f"{width}x{height}+{x}+{y}")

            self.recording_win.deiconify()
            self.recording_win.lift()

            try:
                self.recording_win.attributes("-topmost", True)
            except Exception:
                pass
        except Exception:
            pass


    def hide_recording_overlay(self) -> None:
        #hides recording window
        try: self.recording_win.withdraw()
        except Exception: pass


    def _start_working_animation(self):
        if self._working_animation_after_id is None:
            self._animate_working()

    def _animate_working(self):
        #written by AI
        states = ["working", "working.", "working..", "working..."]
        self._working_dots = (self._working_dots + 1) % len(states)
        self.rpa_status_label.config(text=states[self._working_dots])
        self._working_animation_after_id = self.root.after(500, self._animate_working)

    def _stop_working_animation(self):
        if self._working_animation_after_id is not None:
            self.root.after_cancel(self._working_animation_after_id)
            self._working_animation_after_id = None
            self._working_dots = 0

    def _start_online_animation(self):
        if self._online_animation_after_id is None:
            self._online_pulse_index = 0
            self._animate_online()

    def _animate_online(self):
        # green puls animation
        colors = ["#22C55E", "#16A34A","#000000", "#15803D", "#16A34A"]
        color = colors[self._online_pulse_index]

        self.status_dot.config(fg=color)

        self._online_pulse_index = (self._online_pulse_index + 1) % len(colors)
        self._online_animation_after_id = self.root.after(1000, self._animate_online)

    def _stop_online_animation(self):
        if self._online_animation_after_id is not None:
            self.root.after_cancel(self._online_animation_after_id)
            self._online_animation_after_id = None

    
    def append_log_line(self, log_line) -> None:
        #appends the console-style log
        self.log_text.config(state="normal")
        now = datetime.datetime.now().strftime("%H:%M")
        # self.log_text.insert("end", f"[{now}] {log_line}\n") #activate in PROD
        self.log_text.insert("end", f"{log_line}\n")
        self.log_text.config(state="disabled")
        self.log_text.see("end")


    def do_something(self):
        pass
   

    def shutdown(self) -> Never | None:
        if self._closing: return
        self._closing = True

        try: self.robot_runtime.recording_service.stop()
        except Exception: pass

        self.root.destroy()
        

class RPASimulator:
    """ simulating the behaviour of the external RPA software"""
    #temporary! ignore this class in all eveluations

    def __init__(self):
        print("started")
        #self.run()  #remove if started from main.
        with open("handover.txt", "w", encoding="utf-8") as f:
            f.write("system_state=idle")
        time.sleep(1)


    def check_for_rebootflag(self):
        import os.path
        if os.path.isfile("reboot_ready.flag"):
            os.remove("reboot_ready.flag")
            print("reboot_ready.flag found, rebooting main.py")
            time.sleep(2)
            import subprocess, sys
            #/home/elias/environments/venv/bin/python3 for venv
            subprocess.run([sys.executable, "main.py",], start_new_session=True)


    def append_system_log(self, text: str, job_id=None):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") #+ str(datetime.datetime.now().microsecond)[:-4]
   
        job_part = f" | JOB {job_id}" if job_id else ""
        message = f"{timestamp} | RPA{job_part} | {text} \n"

        for i in range(5):
            try:
                with open("system.log", "a", encoding="utf-8") as f:
                    f.write(message)
                    f.flush()
                return 

            except Exception as e:
                print(f"{i}st warning from append_system_log():", e)
    

    #läs handover.txt och kolla ifall "queued"
    def run(self):
        print("RPASimulator() I'm alive")
        self.append_system_log("RPA I'm alive")

        #do a time-check to not start old jobs? 

        while(1):
            self.check_for_rebootflag()
            time.sleep(1)
            h = {}
            with open("handover.txt", "r", encoding="utf-8") as f:
                for row in f:   
                    row = row.strip()
                    if not row: continue  
                    if "=" not in row: raise ValueError(f"Invalid row in handover: {row}")
                    key, value = row.split("=", 1)
                    h[key.strip()] = value.strip()

            system_state = h.get("system_state")
            job_id = h.get("job_id") 
            job_type = h.get("job_type")

            if system_state != "job_queued":
                continue


            #om "queued ändra till "job_running" och sov 2 sek
            else:
                handover_data = { "system_state": "job_running",  "job_id": job_id }
                self.append_system_log(f"system_state: job_queued -> job_running", job_id)
                self.append_system_log(f"recieved:{handover_data}",job_id)

                print("RPASimulator() jobb påbörjas. system_state: job_queued -> job_running")
                
                dir_path = os.path.dirname(os.path.abspath("handover.txt"))
                fd, temp_path = tempfile.mkstemp(dir=dir_path)

                with os.fdopen(fd, "w", encoding="utf-8") as tmp:
                    for key, value in handover_data.items():
                        if value is None:
                            value = ""
                        tmp.write(f"{key}={value}\n")

                    tmp.flush()
                    os.fsync(tmp.fileno())
                os.replace(temp_path, "handover.txt")

                processtid= random.randint(2,4)
                time.sleep(processtid)
                self.append_system_log(f"screen_1 completed", job_id)
                time.sleep(3)
                self.append_system_log(f"screen_2 completed", job_id)

                #ändra sen till "job_verifying"
                handover_data = { "system_state": "job_verifying",  "job_id": job_id , "job_type": job_type}
            
                dir_path = os.path.dirname(os.path.abspath("handover.txt"))
                fd, temp_path = tempfile.mkstemp(dir=dir_path)

                with os.fdopen(fd, "w", encoding="utf-8") as tmp:
                    for key, value in handover_data.items():
                        if value is None:
                            value = ""
                        tmp.write(f"{key}={value}\n")

                    tmp.flush()
                    os.fsync(tmp.fileno())
                os.replace(temp_path, "handover.txt")
                print("async RPASimulator(): handover, system_state: job_running -> job_verifying")
                self.append_system_log(f"done, system_state: job_running -> job_verifying", job_id)



def main() -> None:
    ui = DashboardUI()
    robot_runtime = RobotRuntime(ui)
    rpa_simulator = RPASimulator()

    ui.attach_runtime(robot_runtime)

    threading.Thread(target=robot_runtime.run, daemon=True).start()
    threading.Thread(target=robot_runtime.poll_for_stop_flag, daemon=True).start() # kills py on manual external-RPA stop by operator (add logic to external RPA)
    threading.Thread(target=rpa_simulator.run, daemon=True).start() #replace with external RPA when deployed

    ui.run()


if __name__ == "__main__":
    main()