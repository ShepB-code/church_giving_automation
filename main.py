import church
import write

def main():
    doc_name = input("What would you like to name your word document?\nInput: ")
    print("Starting Task...")
    date_col = 3 #"C"
    giving_type_col = 5 #"E"
    category_col = 6 #"F"
    amount_col = 7 #"G"
    check_col = 9 #"H"

    #my_dict = church.store_data(date_col, category_col, amount_col, check_col, "2020")
    print("Organizing Data...")
    #church.write_json(my_dict)


    write.write_doc(doc_name)
    print("Writing Data...")
    print("Done!")



if __name__ == "__main__":
    main()