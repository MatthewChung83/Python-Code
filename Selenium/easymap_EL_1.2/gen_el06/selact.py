### 進入網站 勾選同意 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'ok')))
    Weblog = '(ok--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'ok')))
    driver.switch_to.default_content()
    Weblog = '(ok switch frame--Access)'
except:
    Weblog = '(ok--Fail)'
entrylog = entrylog + Weblog

if '(ok--Access)' in entrylog:
    driver.find_element_by_id('ok').click()
elif '(ok switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element(By.ID,'ok').click()
    driver.switch_to.default_content()
else:
    driver.quit()

### 進入網站 同意進入 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'yes')))
    Weblog = '(yes--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'yes')))
    driver.switch_to.default_content()
    Weblog = '(yes switch frame--Access)'
except:
    Weblog = '(yes--Fail)'
entrylog = entrylog + Weblog

if '(yes--Access)' in entrylog:
    driver.find_element_by_id('yes').click()
elif '(yes switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element_by_id('yes').click()
    driver.switch_to.default_content()
else:
    driver.quit()

### 進入網站 登入帳號 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'Img2')))
    Weblog = '(Img2--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'Img2')))
    driver.switch_to.default_content()
    Weblog = '(Img2 switch frame--Access)'
except:
    Weblog = '(Img2--Fail)'
entrylog = entrylog + Weblog

if '(Img2--Access)' in entrylog:
    driver.find_element_by_id('Img2').click()
elif '(Img2 switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element_by_id('Img2').click()
    driver.switch_to.default_content()
else:
    driver.quit()

### 進入網站 輸入帳號 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'aa-uid')))
    Weblog = '(aa-uid--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'aa-uid')))
    driver.switch_to.default_content()
    Weblog = '(aa-uid switch frame--Access)'
except:
    Weblog = '(aa-uid--Fail)'
entrylog = entrylog + Weblog

if '(aa-uid--Access)' in entrylog:
    driver.find_element(By.ID, 'aa-uid').send_keys(acc_name)
elif '(aa-uid switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element(By.ID, 'aa-uid').send_keys(acc_name)
    driver.switch_to.default_content()
else:
    driver.quit()

### 進入網站 輸入密碼 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'aa-passwd')))
    Weblog = '(aa-passwd--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'aa-passwd')))
    driver.switch_to.default_content()
    Weblog = '(aa-passwd switch frame--Access)'
except:
    Weblog = '(aa-passwd--Fail)'
entrylog = entrylog + Weblog

if '(aa-passwd--Access)' in entrylog:
    driver.find_element(By.ID, 'aa-passwd').send_keys(acc_pwd)
elif '(aa-passwd switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element(By.ID, 'aa-passwd').send_keys(acc_pwd)
    driver.switch_to.default_content()
else:
    driver.quit()

### 進入網站 解析captcha ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'AAAIden1')))
    Weblog = '(AAAIden1--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'AAAIden1')))
    driver.switch_to.default_content()
    Weblog = '(AAAIden1 switch frame--Access)'
except:
    Weblog = '(AAAIden1--Fail)'
entrylog = entrylog + Weblog

if '(AAAIden1--Access)' in entrylog:
    img_ele = driver.find_element_by_id('AAAIden1')
    img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
elif '(AAAIden1 switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    img_ele = driver.find_element_by_id('AAAIden1')
    img = Image.open(io.BytesIO(img_ele.screenshot_as_png) ).resize((150,35)).convert('RGB')
    driver.switch_to.default_content()
else:
    driver.quit()

### 進入網站 輸入captcha ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME,'aa-captchaID')))
    Weblog = '(aa-captchaID--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME,'a-captchaID')))
    driver.switch_to.default_content()
    Weblog = '(a-captchaID switch frame--Access)'
except:
    Weblog = '(aa-captchaID--Fail)'
entrylog = entrylog + Weblog

if '(aa-captchaID--Access)' in entrylog:
    captchaID = input('')
    driver.find_element_by_name('aa-captchaID').send_keys(captchaID)
elif '(a-captchaID switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    captchaID = input('')
    driver.find_element_by_name('aa-captchaID').send_keys(captchaID)
    driver.switch_to.default_content()
else:
    driver.quit()
    
### 進入網站 進入網站 ###
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'submit_hn')))
    Weblog = '(submit_hn--Access)'
except selenium.common.exceptions.TimeoutException:
    driver.switch_to.frame('untie')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,'submit_hn')))
    driver.switch_to.default_content()
    Weblog = '(submit_hn switch frame--Access)'
except:
    Weblog = '(submit_hn--Fail)'
entrylog = entrylog + Weblog

if '(submit_hn--Access)' in entrylog:
    driver.find_element(By.ID, 'submit_hn').click()
elif '(submit_hn switch frame--Access)' in entrylog:
    driver.switch_to.frame('untie')
    driver.find_element(By.ID, 'submit_hn').click()
    driver.switch_to.default_content()
else:
    driver.quit()