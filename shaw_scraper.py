from zoneinfo import ZoneInfo
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

import xlwings as xw
import schedule
import time

shaw_theatres = {  # 7 theatres
    # "Lido":"1",  # shit still under maintainance
    "Balestiar":"3",
    # "Jewel":"8",
    # "Paya Lebar":"9",
    # "Waterway Point": "7",
    # "Nex":"5", 
    # "Lot One":"4",
    }

def scrape_data():
    # GET TIME OF PAGE REQUEST
    tz = ZoneInfo("Asia/Singapore")
    dt = datetime.now(tz)
    date = dt.strftime("%d-%m-%y")
    time = dt.strftime("%H:%M:%S")
    print(date, time)

    movie_data = []
    showtime_data = []
    
    # SET SELENIUM DRIVER
    driver = webdriver.Chrome()
    url = "https://shaw.sg/theatre/location/"

    for theatre, id in shaw_theatres.items():
        print(10*"-",theatre, id, 10*"-")

        # RETRIEVE WEB PAGE OF SPECIFIED THEATRE
        driver.get(url+id)

        # CREATE EMPTY SESSION LIST FOR SPECIFIED THEATRE
        session_data = []
        session_dreamers_bales = []
        session_dreamers_jewel = []
        
        try:  
            movie_list = driver.find_elements(By.CSS_SELECTOR, "div.movies_item-movie")  # get movies showing on that date
            for movie in movie_list:

                # FIND MOVIE TITLE
                title = movie.find_element(By.CSS_SELECTOR, "div.title").text
                
                # FIND EARLIEST TIMING, HALL, & SESSION CODE
                info = movie.find_element(By.CSS_SELECTOR, "a.cell")
                showtime = (info.text).replace("*","").replace("+","")
                showtime = "@" + showtime
                hall = (info.get_attribute("data-balloon").split("\n")[0]).title()

                if "Lumiere" in hall or "Premiere" in hall:
                    hall_type = "Premium"
                elif "Dreamers" in hall:
                    hall_type = "Dreamers"
                else:
                    hall_type = "Normal"

                session_no = info.get_attribute("href").replace("https://shaw.sg/seat-selection/", "")
                
                # HANDLING THE FUCKIN DREAMER THEATRES
                if hall_type == "Dreamers":
                    if theatre == "Jewel":
                        session_dreamers_jewel.append(session_no)
                    else:  # theatre == "Balestiar":
                        session_dreamers_bales.append(session_no)
                else:  # non dreamers (normal)
                    session_data.append(session_no)

                # SAVE MOVIE DATA INTO LIST
                movie_data.append([time, theatre, title, str(showtime), hall, hall_type, str(session_no)])
                print(title, showtime, hall, hall_type, session_no)
            
            print("Normal sessions:", session_data)
            print("Dreamer sessions:", session_dreamers_jewel, session_dreamers_bales)
            
            # GET SEAT DATA FROM NORMAL SESSIONS FOR SPECIFIED THEATRE
            for session in session_data: 
                driver.get("https://shaw.sg/seat-selection/"+session)

                try:
                    element = WebDriverWait(driver, 10).until(  # wait until shit loads
                        EC.presence_of_element_located((By.ID, "DiagramTest_canvas_diagramLayer")))
                    seat_overview = driver.find_element(By.ID, "DiagramTest_canvas_diagramLayer")
                    avail_seats = len(seat_overview.find_elements(By.CSS_SELECTOR, "rect.AV"))
                    onhold_seats = len(seat_overview.find_elements(By.CSS_SELECTOR, "rect.OH"))
                    sold_seats = len(seat_overview.find_elements(By.CSS_SELECTOR, "rect.SO"))
                    sold_seats += len(seat_overview.find_elements(By.CSS_SELECTOR, "rect.BL"))
                    total_seats = avail_seats + onhold_seats + sold_seats
                    occu_rate = round(((sold_seats / total_seats)*100), 2)

                    # SAVE SHOWTIME DATA INTO LIST
                    showtime_data.append([str(session), avail_seats, onhold_seats, sold_seats, total_seats, occu_rate])
                    print(f">>> Session No., {session}:, Total Seats: {total_seats}, Avail. Seats: {avail_seats}, Unvail. Seats: {sold_seats}")

                except:
                    pass

            # GET SEAT DATA FROM BALESTAIR DREAMER SESSIONS
            for session in session_dreamers_bales:
                driver.get("https://shaw.sg/seat-selection/"+session)

                try:
                    element = WebDriverWait(driver, 10).until(  # wait until shit loads
                        EC.presence_of_element_located((By.ID, "dreamer-available-seat")))
                    avail_seats = int(driver.find_element(By.ID, "dreamer-available-seat").text)
                    total_seats = 25; onhold_seats = 0
                    sold_seats = total_seats - avail_seats
                    occu_rate = round(((sold_seats / total_seats)*100), 2)

                    # SAVE SHOWTIME DATA INTO LIST
                    showtime_data.append([str(session), avail_seats, onhold_seats, sold_seats, total_seats, occu_rate])
                    print(f">>> Session No., {session}:, Total Seats: {total_seats}, Avail. Seats: {avail_seats}, Unvail. Seats: {sold_seats}")
                
                except:
                    pass

            # GET SEAT DATA FROM JEWEL DREAMER SESSIONS
            for session in session_dreamers_jewel:
                driver.get("https://shaw.sg/seat-selection/"+session)

                try:
                    element = WebDriverWait(driver, 10).until(  # wait until shit loads
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.dreamers-ticket-section")))
                    
                    add_tickets = driver.find_elements(By.CLASS_NAME, "fa.fa-plus.vaccinated-hall-plus.quantity-icons")
                    for ticket in add_tickets: ticket = ticket.click()
                    continue_btn = driver.find_element(By.CLASS_NAME, "btn.btn-primary.ticket-select-dreamers")
                    continue_btn.click()
                    ActionChains(driver).move_to_element(WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, "btn.btn-primary.ticket-select-dreamers-confirm")))).click().perform() 

                    try:
                        element = WebDriverWait(driver, 10).until(  # wait until shit loads
                            EC.presence_of_element_located((By.ID, "DiagramTest_canvas_diagramLayer")))
                        seat_overview = driver.find_element(By.ID, "DiagramTest_canvas_diagramLayer")
                        avail_seats = len(seat_overview.find_elements(By.CSS_SELECTOR, "rect.AV"))
                        onhold_seats = len(seat_overview.find_elements(By.CSS_SELECTOR, "rect.OH"))
                        sold_seats = len(seat_overview.find_elements(By.CSS_SELECTOR, "rect.SO"))
                        sold_seats += len(seat_overview.find_elements(By.CSS_SELECTOR, "rect.BL"))
                        total_seats = avail_seats + onhold_seats + sold_seats
                        occu_rate = round(((sold_seats / total_seats)*100), 2)

                        # SAVE SHOWTIME DATA INTO LIST
                        showtime_data.append([str(session), avail_seats, onhold_seats, sold_seats, total_seats, occu_rate])
                        print(f">>> Session No., {session}:, Total Seats: {total_seats}, Avail. Seats: {avail_seats}, Unvail. Seats: {sold_seats}")

                    except:
                        pass
                except:
                    pass

        except:
            pass

    driver.quit()
    print(movie_data)
    print(showtime_data)

    # PREP EXCEL SHEET
    wb = xw.Book('Shaw_Data.xlsx')

    try:
        ws = wb.sheets[date] # opens existing sheet
    except:
        template_ws = wb.sheets['Template']
        template_ws.api.Copy(Before=template_ws.api) # copy template sheet
        wb.sheets["Template (2)"].api.Name = date # turn copied template sheet into new sheet
        ws = wb.sheets[date]

    # GET TOTAL NO. OF ROWS
    row_num = 2  # ignore header
    print(ws.range("G"+str(row_num)).value)
    while (ws.range("G"+str(row_num)).value != None):
        row_num += 1

    if row_num == 2:
        # ADD MOVIE RECORD TO BLANK SHEET
        for movie in movie_data:
            ws.range('A'+str(row_num)).value = movie
            row_num += 1
    else:
        # CHECK FOR EXISTING MOVIE RECORD
        for movie in movie_data:
            for i in range(2, row_num):
                ws_movie = ws.range("G"+str(i)).value
                ws_movie = str(ws_movie)[:-2]
                if movie[-1] == ws_movie:
                    print(f"Found {movie[2]}!")
                    break
                # ADD MOVIE RECORD IF NOT FOUND
                elif (movie[-1] != ws_movie and i == row_num-1):
                    ws.range('A'+str(row_num)).value = movie
                    print(f"Added {movie[2]}!")
                    row_num += 1
        
    # GET TOTAL NO. OF ROWS AGAIN
    row_num = 2  # ignore header
    while (ws.range("G"+str(row_num)).value != None):
        row_num += 1

    # CHECK EXISTING SESSION NO.
    for showtime in showtime_data:
        for i in range(2, row_num):
            ws_session = ws.range("G"+str(i)).value
            ws_session = str(ws_session)[:-2]
            # ADD SHOWTIME INFO IF FOUND
            if showtime[0] == ws_session:
                ws.range('H'+str(i)).value = showtime[1:]
                print("Updated session", ws_session)

    print(f"Shaw Data Update Completed at {date} {time}!")
    


scrape_data()

# AUTOMATES IT TO RUN EVERY 15 MINS
schedule.every(10).minutes.do(scrape_data)
timer = 10

while True:
    schedule.run_pending()
    print(f"{int(timer)} mins left...")
    timer -= 1
    if timer <= 0:
        print("10 mins passed! Running data scraper...")
        timer += 10
    time.sleep(60)

