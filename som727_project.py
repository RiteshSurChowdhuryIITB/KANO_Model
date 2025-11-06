import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import argparse



def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--path", type=str, required=False, default=os.getcwd())
    parser.add_argument("--start_cols", type=int, required=False, default=1)
    parser.add_argument("--end_cols", type=int, required=False, default=27)
    parser.add_argument("--mapping", dest="mapping", action='store_true', help="Disable mapping")
    args, unknown = parser.parse_known_args()
    return args


if __name__ == "__main__":
    args = parse_args()

    # Path and data load
    path = args.path
    os.chdir(path)
    df = pd.read_excel("Responses.xlsx")


    # Start and end of columns
    start_cols = args.start_cols
    end_cols = args.end_cols
    df = df.iloc[:, start_cols:end_cols]

    # Functional and dysfunctional columns extraction
    columns = df.columns.tolist()
    pairs = [(columns[i], columns[i+1]) for i in range(0, len(columns), 2)]

    # Headers of Functional Questions
    header_pairs = {f"Feature_{idx+1}": pair for idx, pair in enumerate(pairs)}

    # Mapping responses to numbers
    response_map = {
        "I’d be happy about it, but it’s not necessary.": 1,
        "For me, it’s a must-have.": 2,
        "It doesn’t matter to me either way.": 3,
        "I can accept it even if it’s not perfect.": 4,
        "I would not like it and would be unhappy.": 5
    }

    # Kano Evaluation Table
    kano_table = {
        (1,1):"Q", (1,2):"A", (1,3):"A", (1,4):"A", (1,5):"O",
        (2,1):"R", (2,2):"I", (2,3):"I", (2,4):"I", (2,5):"M",
        (3,1):"R", (3,2):"I", (3,3):"I", (3,4):"I", (3,5):"M",
        (4,1):"R", (4,2):"I", (4,3):"I", (4,4):"I", (4,5):"M",
        (5,1):"R", (5,2):"R", (5,3):"R", (5,4):"R", (5,5):"Q"
    }

    # Feature Renaming
    if args.mapping:
        feature_mapping = {
            "How do you feel if the buses are clean?": "Clean buses",
            "How do you feel if the buses have comfortable seating and effective air conditioning or heating?": "Seating and effective air conditioning",
            "How do you feel if the buses use clean energy such as electric or green hydrogen?" : "Runs on Green Energy",
            "How do you feel if buses run frequently with short waiting times?" : "Runs frequently",
            "How do you feel if buses usually have enough vacant seats for passengers?": "Enough vacant seats",
            "How do you feel if extra buses are provided during peak times such as before morning classes, lunch breaks, and special events that bring larger crowds?": "Extra buses during peak time",
            "How do you feel if the buses always arrive and depart on schedule?": "Arrives and departs on schedule",
            "How do you feel if the bus app and digital display boards provide accurate real-time information?": "Providing Accurate real-time information",
            "How do you feel if buses are available even at night (with lower frequency)?": "Available even at night",
            "How do you feel if the buses have ramps and reserved seats for passengers requiring assistance?": "Ramps and reserved seats",
            "How do you feel if bus drivers and staff can communicate in local language(s)?" : "Communicates in local language",
            "How do you feel if bus drivers drive carefully on slopes, turns, and speed breakers?": "Drives carefully",
            "How do you feel if traveling by bus at night feels safe?": "Traveling by bus at night feels safe",
        }
    else:
        feature_mapping = None


    # Grouping features
    results = []
    for feature_name, (func_col, dys_col) in header_pairs.items():
        if feature_mapping:
            feat = feature_mapping[func_col]
        else:
            feat = func_col
        feature_df = pd.DataFrame()
        feature_df["Functional"] = df[func_col].astype(str).map(response_map)
        feature_df["Dysfunctional"] = df[dys_col].astype(str).map(response_map)
        feature_df = feature_df.dropna()
        feature_df["Kano_Category"] = feature_df.apply(
            lambda x: kano_table.get((x["Functional"], x["Dysfunctional"]), "Q"), axis=1
        )
        counts = feature_df["Kano_Category"].value_counts().to_dict()
        row = {
            "Feature": feat,
            "A": counts.get("A", 0),
            "O": counts.get("O", 0),
            "M": counts.get("M", 0),
            "I": counts.get("I", 0),
            "R": counts.get("R", 0),
            "Q": counts.get("Q", 0),
            "Total": counts.get("A", 0) + counts.get("O", 0) + counts.get("M", 0) + counts.get("I", 0) + counts.get("R", 0) + counts.get("Q", 0)
        }
        results.append(row)
    grouped = pd.DataFrame(results).set_index("Feature")


    # Determining Final category
    def determine_final_category(row):
        A, O, M, I, R, Q = row["A"], row["O"], row["M"], row["I"], row["R"], row["Q"]
        if (A + O + M) > (I + R + Q):
            return max([("A", A), ("O", O), ("M", M)], key=lambda x: x[1])[0]
        else:
            return max([("I", I), ("R", R), ("Q", Q)], key=lambda x: x[1])[0]

    grouped["Final_Category"] = grouped.apply(determine_final_category, axis=1)

    # Customer satisfaction and dissatisfaction coeefficients calculation
    grouped["Satisfaction"] = (grouped["A"] + grouped["O"]) / (grouped["A"] + grouped["O"] + grouped["M"] + grouped["I"])
    grouped["Dissatisfaction"] = - (grouped["O"] + grouped["M"]) / (grouped["A"] + grouped["O"] + grouped["M"] + grouped["I"])

    print(grouped)

    grouped.to_excel("kano_results.xlsx")
    print("Kano analysis complete!")
    print("Saved results to kano_results.xlsx.")
    
    # plt.figure(figsize=(7,6))
    # plt.scatter(grouped["Dissatisfaction"], grouped["Satisfaction"])
    # for i, txt in enumerate(grouped.index):
    #     plt.text(grouped["Dissatisfaction"].iloc[i]+0.01, grouped["Satisfaction"].iloc[i], txt, fontsize=8)
    # plt.axhline(0, color='gray', linestyle='--')
    # plt.axvline(0, color='gray', linestyle='--')
    # plt.xlabel("Dissatisfaction")
    # plt.ylabel("Satisfaction")
    # plt.title("Customer Satisfaction Coefficient Diagram")
    # plt.grid(True)
    # plt.savefig("customer_satisfaction_diagram.png", dpi=300, bbox_inches='tight')
    # plt.show()

    # plt.figure(figsize=(7,6))
    # plt.scatter(grouped["Satisfaction"],grouped["Dissatisfaction"])
    # for i, txt in enumerate(grouped.index):
    #     plt.text(grouped["Satisfaction"].iloc[i], grouped["Dissatisfaction"].iloc[i]-0.02, txt, fontsize=8)
    # plt.axhline(-0.5, color='gray', linestyle='--')
    # plt.axvline(0.5, color='gray', linestyle='--')
    # plt.xlabel("Satisfaction")
    # plt.ylabel("Dissatisfaction")
    # plt.title("Customer Satisfaction Coefficient Diagram")
    # plt.grid(True)
    # plt.savefig("customer_satisfaction_diagram.png", dpi=300, bbox_inches='tight')
    # plt.show()

    from adjustText import adjust_text
    import matplotlib.pyplot as plt

    plt.figure(figsize=(7,6))
    plt.scatter(grouped["Satisfaction"], grouped["Dissatisfaction"])

    texts = []
    for i, txt in enumerate(grouped.index):
        texts.append(
            plt.text(
                grouped["Satisfaction"].iloc[i],
                grouped["Dissatisfaction"].iloc[i],
                txt,
                fontsize=8
            )
        )

    # Automatically adjust to avoid overlap
    adjust_text(texts, arrowprops=dict(arrowstyle="->", color='gray', lw=0.5))

    plt.axhline(-0.5, color='gray', linestyle='--')
    plt.axvline(0.5, color='gray', linestyle='--')
    plt.xlabel("Satisfaction")
    plt.ylabel("Dissatisfaction")
    plt.title("Customer Satisfaction Coefficient Diagram")
    plt.grid(True)
    plt.tight_layout()
    plt.savefig("customer_satisfaction_diagram.png", dpi=300, bbox_inches='tight')
    plt.show()
