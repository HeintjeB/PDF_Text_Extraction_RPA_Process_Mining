import pandas as pd
import pm4py
import os

file_path = os.path.dirname(__file__)
os.chdir(file_path)

process_flow_df = pd.read_parquet(
    f"{file_path}/fake_confirmations/snappy/process_mining_order_confirmation.snappy.parquet",
    engine="pyarrow",
)

event_log = pm4py.format_dataframe(
    process_flow_df,
    case_id="case",
    activity_key="activity",
    timestamp_key="timestamp",
    timest_format="%Y-%m-%d",
)

net = pm4py.discover_heuristics_net(event_log)
pm4py.view_heuristics_net(net)
