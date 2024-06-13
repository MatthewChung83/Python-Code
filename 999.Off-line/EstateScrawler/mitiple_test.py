import multiprocessing

def worker(file):
    # your subprocess code


if __name__ == '__main__':
    files = ["C:\Py_Project\project\EstateScrawler\test1.py","C:\Py_Project\project\EstateScrawler\test2.py","C:\Py_Project\project\EstateScrawler\test3.py"]
    for i in files:
        p = multiprocessing.Process(target=worker, args=(i,))
        p.start()