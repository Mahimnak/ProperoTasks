from RPA.Robocorp.WorkItems import WorkItems

def fetch_workitems() -> dict:
    """
    Retrieves the input parameters required for the entire process.
    """
    # work_items = WorkItems()
    # work_items.get_input_work_item()
    # work_item = work_items.get_work_item_variables()
    
    work_item = {
            "search phrase":"Football",
            "category":"All",
            "timespan":"1 month"
    }

    return work_item