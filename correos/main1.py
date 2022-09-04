import pandas as pd

def main():

    data = pd.read_csv("nba.csv")
    frame = pd.DataFrame(data)
    print(frame.infer_objects().empty)


    
if __name__ == '__main__':
    main()

