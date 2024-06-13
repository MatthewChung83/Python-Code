def CAPTCHAOCR(Apurl,imgf,imgp):
    import requests
    url = Apurl
    
    files = {
        "image_file":(imgf,open(imgp,"rb"),'image/jpeg')
    }
    response = requests.request("POST", url, files=files)
    return response.text