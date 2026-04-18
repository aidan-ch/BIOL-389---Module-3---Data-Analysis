import pandas as pd
import matplotlib.pyplot as plt
import os

data_path = "Module 3 clampfit data.xlsx"

# per recording numbers
num_sweeps = 12
num_pulses = 4


def initialize_directories():
    if not os.path.exists("Plots"):
        os.mkdir("Plots")
    if not os.path.exists("Data"):
        os.mkdir("Data")


class Recording:
    def __init__(self, recording_df):
        self.recording_df = recording_df
        self.sweeps = {}

        for sweep_num in range(1, num_sweeps + 1):
            sweep_row_mask = recording_df["Sweep"] == sweep_num
            sweep_df = recording_df[sweep_row_mask]
            self.sweeps[sweep_num] = Sweep(sweep_df)

    def get_sweep(self, sweep_num):
        return self.sweeps[sweep_num]


class Sweep:
    def __init__(self, sweep_df):
        self.sweep_df = sweep_df
        self.sweep_num = sweep_df["Sweep"].iloc[0]

        # pulse_num : PD series of pulse data (same cols as recording data but no pulse #, sweep #)
        self.pulses = {}

        for pulse_row, row in self.sweep_df.iterrows():
            # add the entire row but first 2 columns (pulse number, sweep number)
            pulse_num = row["Pulse"]
            self.pulses[pulse_num] = Pulse(row["Volley peak (V)":], pulse_num)

    def get_pulse(self, pulse_num):
        return self.pulses[pulse_num]


class Pulse:
    def __init__(self, data_series, pulse_num):
        self.data_series = data_series
        self.pulse_num = pulse_num

        self.volley_peak = data_series["Volley peak (V)"]
        self.volley_min = data_series["Volley min (V)"]
        self.volley_peak_time = data_series["Volley Peak (ms)"]
        self.volley_min_time = data_series["Volley min (ms)"]
        self.epsp = data_series["EPSP (V)"]
        self.epsp_time = data_series["EPSP (ms)"]
        self.pulse_end_time = data_series["Pulse end time (ms)"]

    def get_volley_peak(self):
        return self.volley_peak

    def get_volley_min(self):
        return self.volley_min

    def get_volley_amplitude(self):
        return abs(self.volley_peak - self.volley_min)

    def get_epsp_amplitude(self):
        return abs(self.epsp)

    def get_volley_peak_time(self):
        return self.volley_peak_time

    def get_volley_min_time(self):
        return self.volley_min_time

    # (time to peak + time to min) / 2  - pulse_end_time
    def get_time_to_volley(self):
        return (
                    self.volley_peak_time + self.volley_min_time) / 2 - self.pulse_end_time

    def get_pulse_end_time(self):
        return self.pulse_end_time

    def get_epsp(self):
        return self.epsp

    def get_epsp_time(self):
        return self.epsp_time


# FIBER VOLLEY AMPLITUDE = PEAK - MIN
'''
Returns a pandas DataFrame:

Columns: Pulse, Sweep, avg FV amplitude (mV)
    --> avg FV amplitude is across all recordings (in millivolts)
'''


def get_avg_fv_amplitude(recordings):
    data_dict = {"Pulse": [], "Sweep": [], "avg FV amplitude (mV)": []}
    num_recordings = len(recordings)

    for sweep_num in range(1, num_sweeps + 1):
        for pulse_num in range(1, num_pulses + 1):
            sum_FV_amplitude = 0
            for recording in recordings:
                sweep = recording.get_sweep(sweep_num)
                pulse = sweep.get_pulse(pulse_num)
                # multiply by 1000 to get to millivolts
                sum_FV_amplitude += pulse.get_volley_amplitude() * 1000

            data_dict["Pulse"].append(pulse_num)
            data_dict["Sweep"].append(sweep_num)
            avg_fv_amplitude = sum_FV_amplitude / num_recordings
            data_dict["avg FV amplitude (mV)"].append(avg_fv_amplitude)

    return pd.DataFrame(data_dict)


# generates + plots the data and saves it to a CSV
def generate_pulse_fv_amplitude_vs_sweep(recordings):
    fig, ax = plt.subplots(1, 1)
    fig.suptitle("Average Fiber Volley Amplitude vs. Sweep")
    avg_fv_amplitude_df = get_avg_fv_amplitude(recordings)

    # save the data to a .csv
    avg_fv_amplitude_df.to_csv(
        "Data/Average Fiber Volley Amplitude vs. Sweep.csv",
        index=False)

    df_pulse_1 = avg_fv_amplitude_df[avg_fv_amplitude_df["Pulse"] == 1]
    df_pulse_2 = avg_fv_amplitude_df[avg_fv_amplitude_df["Pulse"] == 2]
    df_pulse_3 = avg_fv_amplitude_df[avg_fv_amplitude_df["Pulse"] == 3]
    df_pulse_4 = avg_fv_amplitude_df[avg_fv_amplitude_df["Pulse"] == 4]
    x_values = df_pulse_1["Sweep"]

    ax.plot(x_values, df_pulse_1["avg FV amplitude (mV)"], label="Pulse 1")
    ax.plot(x_values, df_pulse_2["avg FV amplitude (mV)"], label="Pulse 2")
    ax.plot(x_values, df_pulse_3["avg FV amplitude (mV)"], label="Pulse 3")
    ax.plot(x_values, df_pulse_4["avg FV amplitude (mV)"], label="Pulse 4")
    ax.set(xlabel="Sweep Number", ylabel="Average FV amplitude (mV)")
    fig.legend()
    fig.savefig("Plots/Pulse FV Amplitude vs Sweep Number")
    fig.show()
    plt.close(fig)


# DEFINED TIME TO VOLLEY AS (TIME TO PEAK + TIME TO MIN)/2 - pulse end time
def get_avg_time_to_volley(recordings):
    data_dict = {"Pulse": [], "Sweep": [], "Average time to FV (ms)": []}
    for sweep_num in range(1, num_sweeps + 1):
        for pulse_num in range(1, num_pulses + 1):
            sum_time_to_volley = 0
            for recording in recordings:
                sweep = recording.get_sweep(sweep_num)
                pulse = sweep.get_pulse(pulse_num)

                sum_time_to_volley += pulse.get_time_to_volley()

            data_dict["Pulse"].append(pulse_num)
            data_dict["Sweep"].append(sweep_num)
            data_dict["Average time to FV (ms)"].append(
                sum_time_to_volley / len(recordings))

    return pd.DataFrame(data_dict)


# generate plots for average time to volley peak (from pulse end) vs. Sweep number
def generate_pulse_fv_time_vs_sweep(recordings):
    fig, ax = plt.subplots(1, 1)
    fig.suptitle("Average Time to Volley vs. Sweep")
    avg_time_to_volley = get_avg_time_to_volley(recordings)

    # save the data to a .csv
    avg_time_to_volley.to_csv(
        "Data/Average time to FV (ms) vs. Sweep.csv",
        index=False)

    df_pulse_1 = avg_time_to_volley[avg_time_to_volley["Pulse"] == 1]
    df_pulse_2 = avg_time_to_volley[avg_time_to_volley["Pulse"] == 2]
    df_pulse_3 = avg_time_to_volley[avg_time_to_volley["Pulse"] == 3]
    df_pulse_4 = avg_time_to_volley[avg_time_to_volley["Pulse"] == 4]
    x_values = df_pulse_1["Sweep"]

    ax.plot(x_values, df_pulse_1["Average time to FV (ms)"],
            label="Pulse 1")
    ax.plot(x_values, df_pulse_2["Average time to FV (ms)"],
            label="Pulse 2")
    ax.plot(x_values, df_pulse_3["Average time to FV (ms)"],
            label="Pulse 3")
    ax.plot(x_values, df_pulse_4["Average time to FV (ms)"],
            label="Pulse 4")
    ax.set(xlabel="Sweep Number", ylabel="Average time to volley (ms)")
    fig.legend()
    fig.savefig("Plots/Time to volley vs Sweep Number")
    fig.show()
    plt.close(fig)


'''
Returns a pandas DataFrame:

Columns: Pulse, Sweep, avg EPSP amplitude (mV)
    --> avg FV amplitude is across all recordings (in millivolts)
'''


def get_avg_epsp_amplitude(recordings):
    data_dict = {"Pulse": [], "Sweep": [], "avg EPSP amplitude (mV)": []}
    num_recordings = len(recordings)

    for sweep_num in range(1, num_sweeps + 1):
        for pulse_num in range(1, num_pulses + 1):
            sum_epsp_amplitude = 0
            for recording in recordings:
                sweep = recording.get_sweep(sweep_num)
                pulse = sweep.get_pulse(pulse_num)
                # multiply by 1000 to get to millivolts
                sum_epsp_amplitude += pulse.get_epsp_amplitude() * 1000

            data_dict["Pulse"].append(pulse_num)
            data_dict["Sweep"].append(sweep_num)

            avg_epsp_amplitude = sum_epsp_amplitude / num_recordings
            data_dict["avg EPSP amplitude (mV)"].append(avg_epsp_amplitude)

    return pd.DataFrame(data_dict)


# plot average EPSP amplitude (per pulse) vs. sweep number
def generate_avg_epsp_vs_sweep(recordings):
    avg_epsp_df = get_avg_epsp_amplitude(recordings)

    fig, ax = plt.subplots(1, 1)
    fig.suptitle("Average EPSP Amplitude vs. Sweep")

    # save the data to a .csv
    avg_epsp_df.to_csv(
        "Data/Average EPSP Amplitude vs. Sweep.csv", index=False)

    df_pulse_1 = avg_epsp_df[avg_epsp_df["Pulse"] == 1]
    df_pulse_2 = avg_epsp_df[avg_epsp_df["Pulse"] == 2]
    df_pulse_3 = avg_epsp_df[avg_epsp_df["Pulse"] == 3]
    df_pulse_4 = avg_epsp_df[avg_epsp_df["Pulse"] == 4]
    x_values = df_pulse_1["Sweep"]

    ax.plot(x_values, df_pulse_1["avg EPSP amplitude (mV)"], label="Pulse 1")
    ax.plot(x_values, df_pulse_2["avg EPSP amplitude (mV)"], label="Pulse 2")
    ax.plot(x_values, df_pulse_3["avg EPSP amplitude (mV)"], label="Pulse 3")
    ax.plot(x_values, df_pulse_4["avg EPSP amplitude (mV)"], label="Pulse 4")
    ax.set(xlabel="Sweep Number", ylabel="Average EPSP amplitude (mV)")
    fig.legend()
    fig.savefig("Plots/Pulse EPSP Amplitude vs Sweep Number")
    fig.show()
    plt.close(fig)


'''
Returns a pandas DataFrame:

Columns: Pulse, Sweep, avg EPSP/Fiber volley amplitude
'''


# calculates the average of ratios not the ratio of averages
def get_avg_epsp_amplitude_fv_normalized(recordings):
    data_dict = {"Pulse": [], "Sweep": [],
                 "Average(EPSP amplitude/Fiber volley amplitude)": []}
    for sweep_num in range(1, num_sweeps + 1):
        for pulse_num in range(1, num_pulses + 1):
            sum_epsp_over_fv_amplitude = 0
            for recording in recordings:
                sweep = recording.get_sweep(sweep_num)
                pulse = sweep.get_pulse(pulse_num)
                epsp_amplitude = pulse.get_epsp_amplitude()
                fv_amplitude = pulse.get_volley_amplitude()
                sum_epsp_over_fv_amplitude += epsp_amplitude / fv_amplitude

            data_dict["Pulse"].append(pulse_num)
            data_dict["Sweep"].append(sweep_num)
            data_dict["Average(EPSP amplitude/Fiber volley amplitude)"].append(
                sum_epsp_over_fv_amplitude / len(recordings))

    return pd.DataFrame(data_dict)


# plot average(EPSP/Fiber volley amplitude) per pulse vs. Sweep
# EPSP/FV is a normalized synaptic response metric; changes relative to pulse 1 suggest facilitation or depression
def generate_avg_epsp_vs_pulse_fv_normalized(recordings):
    df = get_avg_epsp_amplitude_fv_normalized(recordings)

    fig, ax = plt.subplots(1, 1)
    fig.suptitle("Synaptic Efficacy Across Sweeps")

    # save the data to a .csv
    df.to_csv(
        "Data/Average Synaptic Efficacy vs Sweep.csv",
        index=False)

    df_pulse_1 = df[df["Pulse"] == 1]
    df_pulse_2 = df[df["Pulse"] == 2]
    df_pulse_3 = df[df["Pulse"] == 3]
    df_pulse_4 = df[df["Pulse"] == 4]
    x_values = df_pulse_1["Sweep"]

    ax.plot(x_values,
            df_pulse_1["Average(EPSP amplitude/Fiber volley amplitude)"],
            label="Pulse 1")
    ax.plot(x_values,
            df_pulse_2["Average(EPSP amplitude/Fiber volley amplitude)"],
            label="Pulse 2")
    ax.plot(x_values,
            df_pulse_3["Average(EPSP amplitude/Fiber volley amplitude)"],
            label="Pulse 3")
    ax.plot(x_values,
            df_pulse_4["Average(EPSP amplitude/Fiber volley amplitude)"],
            label="Pulse 4")
    ax.set(xlabel="Sweep Number",
           ylabel="Average Synaptic Efficacy")
    fig.legend()
    fig.savefig("Plots/Average Synaptic Efficacy vs Sweep")
    fig.show()
    plt.close(fig)


# plot avg(EPSP/FV)_pulse2/avg(EPSP/FV)_pulse1, avg(EPSP/FV)_pulse3/avg(EPSP/FV)_pulse1
# avg(EPSP/FV)_pulse4/avg(EPSP/FV)_pulse1
def generate_epspToFV_ratios_pulse_comparisons(recordings):
    data_dict = {
        "Sweep": [],
        "Pulse 2/Pulse 1": [],
        "Pulse 3/Pulse 1": [],
        "Pulse 4/Pulse 1": [],
    }

    for sweep_num in range(1, num_sweeps + 1):
        pulse_2_over_1_values = []
        pulse_3_over_1_values = []
        pulse_4_over_1_values = []

        for recording in recordings:
            sweep = recording.get_sweep(sweep_num)

            pulse_1 = sweep.get_pulse(1)
            pulse_2 = sweep.get_pulse(2)
            pulse_3 = sweep.get_pulse(3)
            pulse_4 = sweep.get_pulse(4)

            fv_1 = pulse_1.get_volley_amplitude()
            fv_2 = pulse_2.get_volley_amplitude()
            fv_3 = pulse_3.get_volley_amplitude()
            fv_4 = pulse_4.get_volley_amplitude()

            epsp_1 = pulse_1.get_epsp_amplitude()
            epsp_2 = pulse_2.get_epsp_amplitude()
            epsp_3 = pulse_3.get_epsp_amplitude()
            epsp_4 = pulse_4.get_epsp_amplitude()

            # Skip this recording for this sweep if any denominator is zero.
            if fv_1 == 0 or fv_2 == 0 or fv_3 == 0 or fv_4 == 0:
                continue

            pulse_1_ratio = epsp_1 / fv_1
            pulse_2_ratio = epsp_2 / fv_2
            pulse_3_ratio = epsp_3 / fv_3
            pulse_4_ratio = epsp_4 / fv_4

            # Skip if pulse 1 normalized response is zero.
            if pulse_1_ratio == 0:
                continue

            pulse_2_over_1_values.append(pulse_2_ratio / pulse_1_ratio)
            pulse_3_over_1_values.append(pulse_3_ratio / pulse_1_ratio)
            pulse_4_over_1_values.append(pulse_4_ratio / pulse_1_ratio)

        data_dict["Sweep"].append(sweep_num)
        data_dict["Pulse 2/Pulse 1"].append(
            sum(pulse_2_over_1_values) / len(pulse_2_over_1_values)
            if pulse_2_over_1_values else float("nan")
        )
        data_dict["Pulse 3/Pulse 1"].append(
            sum(pulse_3_over_1_values) / len(pulse_3_over_1_values)
            if pulse_3_over_1_values else float("nan")
        )
        data_dict["Pulse 4/Pulse 1"].append(
            sum(pulse_4_over_1_values) / len(pulse_4_over_1_values)
            if pulse_4_over_1_values else float("nan")
        )

    df = pd.DataFrame(data_dict)

    df.to_csv(
        "Data/Average Synaptic Efficacy between pulses vs Sweep.csv",
        index=False,
    )

    fig, ax = plt.subplots(1, 1, layout = "constrained")
    fig.suptitle("Synaptic Efficacy Across Pulses")

    ax.plot(df["Sweep"], df["Pulse 2/Pulse 1"], label="Pulse 2 / Pulse 1")
    ax.plot(df["Sweep"], df["Pulse 3/Pulse 1"], label="Pulse 3 / Pulse 1")
    ax.plot(df["Sweep"], df["Pulse 4/Pulse 1"], label="Pulse 4 / Pulse 1")

    ax.set(
        xlabel="Sweep Number",
        ylabel="Average (SE Pulse n) / (SE Pulse 1)",
    )
    fig.legend()
    fig.savefig(
        "Plots/Average Synaptic Efficacy (pulse comparisons) vs Sweep.png"
    )
    fig.show()
    plt.close(fig)


def summarize_facilitation_across_sweeps(recordings):
    data_dict = {
        "Sweep": [],
        "Pulse 2/Pulse 1": [],
        "Pulse 3/Pulse 1": [],
        "Pulse 4/Pulse 1": [],
    }

    for sweep_num in range(1, num_sweeps + 1):
        pulse_2_over_1_values = []
        pulse_3_over_1_values = []
        pulse_4_over_1_values = []

        for recording in recordings:
            sweep = recording.get_sweep(sweep_num)

            pulse_1 = sweep.get_pulse(1)
            pulse_2 = sweep.get_pulse(2)
            pulse_3 = sweep.get_pulse(3)
            pulse_4 = sweep.get_pulse(4)

            fv_1 = pulse_1.get_volley_amplitude()
            fv_2 = pulse_2.get_volley_amplitude()
            fv_3 = pulse_3.get_volley_amplitude()
            fv_4 = pulse_4.get_volley_amplitude()

            epsp_1 = pulse_1.get_epsp_amplitude()
            epsp_2 = pulse_2.get_epsp_amplitude()
            epsp_3 = pulse_3.get_epsp_amplitude()
            epsp_4 = pulse_4.get_epsp_amplitude()

            pulse_1_ratio = epsp_1 / fv_1
            pulse_2_ratio = epsp_2 / fv_2
            pulse_3_ratio = epsp_3 / fv_3
            pulse_4_ratio = epsp_4 / fv_4

            if pulse_1_ratio == 0:
                continue

            pulse_2_over_1_values.append(pulse_2_ratio / pulse_1_ratio)
            pulse_3_over_1_values.append(pulse_3_ratio / pulse_1_ratio)
            pulse_4_over_1_values.append(pulse_4_ratio / pulse_1_ratio)

        data_dict["Sweep"].append(sweep_num)
        data_dict["Pulse 2/Pulse 1"].append(
            sum(pulse_2_over_1_values) / len(pulse_2_over_1_values)
            if pulse_2_over_1_values else float("nan")
        )
        data_dict["Pulse 3/Pulse 1"].append(
            sum(pulse_3_over_1_values) / len(pulse_3_over_1_values)
            if pulse_3_over_1_values else float("nan")
        )
        data_dict["Pulse 4/Pulse 1"].append(
            sum(pulse_4_over_1_values) / len(pulse_4_over_1_values)
            if pulse_4_over_1_values else float("nan")
        )

    facilitation_df = pd.DataFrame(data_dict)

    summary_rows = []
    comparison_columns = [
        "Pulse 2/Pulse 1",
        "Pulse 3/Pulse 1",
        "Pulse 4/Pulse 1",
    ]

    for column in comparison_columns:
        values = facilitation_df[column].dropna()

        if len(values) == 0:
            summary_rows.append({
                "Comparison": column,
                "Mean": float("nan"),
                "Median": float("nan"),
                "Min": float("nan"),
                "Max": float("nan"),
                "SD": float("nan"),
                "Sweeps with facilitation (%)": float("nan"),
                "Sweeps with facilitation (n)": 0,
                "Total sweeps used": 0,
            })
            continue

        summary_rows.append({
            "Comparison": column,
            "Mean": values.mean(),
            "Median": values.median(),
            "Min": values.min(),
            "Max": values.max(),
            "SD": values.std() if len(values) > 1 else 0.0,
            "Sweeps with facilitation (%)": (values.gt(1).sum() / len(
                values)) * 100,
            "Sweeps with facilitation (n)": int(values.gt(1).sum()),
            "Total sweeps used": int(len(values)),
        })

    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv("Data/Facilitation summary across sweeps.csv",
                      index=False)

    return facilitation_df, summary_df


def get_normalized_avg_epsp_amplitude(recordings):
    epsp_df = get_avg_epsp_amplitude(recordings).copy()
    normalized_epsp_dict = {
        "Sweep": [],
        "Pulse": [],
        "Normalized EPSP magnitude (%)": [],
    }

    for pulse in range(1, num_pulses + 1):
        pulse_df = epsp_df[epsp_df["Pulse"] == pulse]
        sweep_1_magnitude = pulse_df.loc[
            pulse_df["Sweep"] == 1, "avg EPSP amplitude (mV)"].iloc[0]

        normalized_epsp_dict["Normalized EPSP magnitude (%)"].extend(
            (pulse_df["avg EPSP amplitude (mV)"] / sweep_1_magnitude * 100).tolist())

        normalized_epsp_dict["Sweep"].extend(range(1, num_sweeps + 1))
        normalized_epsp_dict["Pulse"].extend([pulse] * num_sweeps)

    return pd.DataFrame(normalized_epsp_dict)


# sweep 1 normalized
def generate_normalized_epsp_vs_sweep(recordings):
    normalized_epsp_df = get_normalized_avg_epsp_amplitude(recordings)

    fig, ax = plt.subplots(1, 1)
    fig.suptitle("Normalized EPSP Amplitude vs. Sweep")

    # save the data to a .csv
    normalized_epsp_df.to_csv(
        "Data/Normalized EPSP Amplitude vs. Sweep.csv", index=False)

    df_pulse_1 = normalized_epsp_df[normalized_epsp_df["Pulse"] == 1]
    df_pulse_2 = normalized_epsp_df[normalized_epsp_df["Pulse"] == 2]
    df_pulse_3 = normalized_epsp_df[normalized_epsp_df["Pulse"] == 3]
    df_pulse_4 = normalized_epsp_df[normalized_epsp_df["Pulse"] == 4]
    x_values = df_pulse_1["Sweep"]

    ax.plot(x_values, df_pulse_1["Normalized EPSP magnitude (%)"],
            label="Pulse 1")
    ax.plot(x_values, df_pulse_2["Normalized EPSP magnitude (%)"],
            label="Pulse 2")
    ax.plot(x_values, df_pulse_3["Normalized EPSP magnitude (%)"],
            label="Pulse 3")
    ax.plot(x_values, df_pulse_4["Normalized EPSP magnitude (%)"],
            label="Pulse 4")
    ax.set(xlabel="Sweep Number", ylabel="Normalized EPSP magnitude (% of sweep 1)")
    fig.legend()
    fig.savefig("Plots/Normalized EPSP Amplitude vs Sweep Number")
    fig.show()
    plt.close(fig)

def get_normalized_avg_fv_amplitude(recordings):
    fv_df = get_avg_fv_amplitude(recordings).copy()
    normalized_fv_dict = {
        "Sweep": [],
        "Pulse": [],
        "Normalized FV magnitude (%)": [],
    }

    for pulse in range(1, num_pulses + 1):
        pulse_df = fv_df[fv_df["Pulse"] == pulse]
        sweep_1_magnitude = pulse_df.loc[
            pulse_df["Sweep"] == 1, "avg FV amplitude (mV)"].iloc[0]

        normalized_fv_dict["Normalized FV magnitude (%)"].extend(
            (pulse_df["avg FV amplitude (mV)"] / sweep_1_magnitude * 100).tolist())

        normalized_fv_dict["Sweep"].extend(range(1, num_sweeps + 1))
        normalized_fv_dict["Pulse"].extend([pulse] * num_sweeps)

    return pd.DataFrame(normalized_fv_dict)


# sweep 1 normalized
def generate_normalized_fv_vs_sweep(recordings):
    normalized_fv_df = get_normalized_avg_fv_amplitude(recordings)

    fig, ax = plt.subplots(1, 1)
    fig.suptitle("Normalized FV Amplitude vs. Sweep")

    # save the data to a .csv
    normalized_fv_df.to_csv(
        "Data/Normalized FV Amplitude vs. Sweep.csv", index=False)

    df_pulse_1 = normalized_fv_df[normalized_fv_df["Pulse"] == 1]
    df_pulse_2 = normalized_fv_df[normalized_fv_df["Pulse"] == 2]
    df_pulse_3 = normalized_fv_df[normalized_fv_df["Pulse"] == 3]
    df_pulse_4 = normalized_fv_df[normalized_fv_df["Pulse"] == 4]
    x_values = df_pulse_1["Sweep"]

    ax.plot(x_values, df_pulse_1["Normalized FV magnitude (%)"],
            label="Pulse 1")
    ax.plot(x_values, df_pulse_2["Normalized FV magnitude (%)"],
            label="Pulse 2")
    ax.plot(x_values, df_pulse_3["Normalized FV magnitude (%)"],
            label="Pulse 3")
    ax.plot(x_values, df_pulse_4["Normalized FV magnitude (%)"],
            label="Pulse 4")
    ax.set(xlabel="Sweep Number", ylabel="Normalized FV magnitude (% of sweep 1)")
    fig.legend()
    fig.savefig("Plots/Normalized FV Amplitude vs Sweep Number")
    fig.show()
    plt.close(fig)

# returns a dataframe showing for each sweep, the average drop in fv amplitude
    # from pulse 1 to 2, pulse 2 to 3, pulse 3 to 4
def get_average_drop_fv_amplitude(recordings):
    data_dict = {
        "Sweep": [],
        "Pulse 1 to 2 delta V": [],
        "Pulse 2 to 3 delta V": [],
        "Pulse 3 to 4 delta V": [],
    }

    for sweep in range(1, num_sweeps + 1):
        sum_drop_1_to_2 = 0
        sum_drop_2_to_3 = 0
        sum_drop_3_to_4 = 0
        for recording in recordings:
            sweep_df = recording.get_sweep(sweep)
            pulse_1 = sweep_df.get_pulse(1)
            pulse_2 = sweep_df.get_pulse(2)
            pulse_3 = sweep_df.get_pulse(3)
            pulse_4 = sweep_df.get_pulse(4)

            sum_drop_1_to_2 = pulse_2.get_volley_amplitude() - pulse_1.get_volley_amplitude()
            sum_drop_2_to_3 = pulse_3.get_volley_amplitude() - pulse_2.get_volley_amplitude()
            sum_drop_3_to_4 = pulse_4.get_volley_amplitude() - pulse_3.get_volley_amplitude()

        data_dict["Sweep"].append(sweep)
        data_dict["Pulse 1 to 2 delta V"].append(sum_drop_1_to_2/len(recordings))
        data_dict["Pulse 2 to 3 delta V"].append(sum_drop_2_to_3/len(recordings))
        data_dict["Pulse 3 to 4 delta V"].append(sum_drop_3_to_4/len(recordings))

    return pd.DataFrame(data_dict)

def generate_avg_drop_fv_amplitude(recordings):
    avg_drop_df = get_average_drop_fv_amplitude(recordings)

    fig, ax = plt.subplots(1, 1, layout = "constrained")
    fig.suptitle("Fiber Volley ∆V across pulses vs. Sweep")

    # save the data to a .csv
    avg_drop_df.to_csv(
        "Data/Fiber Volley ∆V across pulses vs. Sweep.csv", index=False)


    x_values = avg_drop_df["Sweep"]

    ax.plot(x_values, avg_drop_df["Pulse 1 to 2 delta V"],
            label="Pulse 2 - Pulse 1")
    ax.plot(x_values, avg_drop_df["Pulse 2 to 3 delta V"],
            label="Pulse 3 - Pulse 2")
    ax.plot(x_values, avg_drop_df["Pulse 3 to 4 delta V"],
            label="Pulse 4 - Pulse 3")

    ax.set(xlabel="Sweep Number", ylabel="Fiber Volley ∆V (V)")
    fig.legend()
    fig.savefig("Plots/Fiber Volley ∆V vs Sweep Number")
    fig.show()
    plt.close(fig)

# returns a dataframe showing for each sweep, the average drop in epsp amplitude
# from pulse 1 to 2, pulse 2 to 3, pulse 3 to 4
def get_average_drop_epsp_amplitude(recordings):
    data_dict = {
        "Sweep": [],
        "Pulse 1 to 2 delta V": [],
        "Pulse 2 to 3 delta V": [],
        "Pulse 3 to 4 delta V": [],
    }

    for sweep in range(1, num_sweeps + 1):
        sum_drop_1_to_2 = 0
        sum_drop_2_to_3 = 0
        sum_drop_3_to_4 = 0
        for recording in recordings:
            sweep_df = recording.get_sweep(sweep)
            pulse_1 = sweep_df.get_pulse(1)
            pulse_2 = sweep_df.get_pulse(2)
            pulse_3 = sweep_df.get_pulse(3)
            pulse_4 = sweep_df.get_pulse(4)

            sum_drop_1_to_2 = pulse_2.get_epsp_amplitude() - pulse_1.get_epsp_amplitude()
            sum_drop_2_to_3 = pulse_3.get_epsp_amplitude() - pulse_2.get_epsp_amplitude()
            sum_drop_3_to_4 = pulse_4.get_epsp_amplitude() - pulse_3.get_epsp_amplitude()

        data_dict["Sweep"].append(sweep)
        data_dict["Pulse 1 to 2 delta V"].append(sum_drop_1_to_2/len(recordings))
        data_dict["Pulse 2 to 3 delta V"].append(sum_drop_2_to_3/len(recordings))
        data_dict["Pulse 3 to 4 delta V"].append(sum_drop_3_to_4/len(recordings))

    return pd.DataFrame(data_dict)

def generate_avg_drop_epsp_amplitude(recordings):
    avg_drop_df = get_average_drop_epsp_amplitude(recordings)

    fig, ax = plt.subplots(1, 1, layout = "constrained")
    fig.suptitle("fEPSP ∆V across pulses vs. Sweep")

    # save the data to a .csv
    avg_drop_df.to_csv(
        "Data/fEPSP ∆V across pulses vs. Sweep.csv", index=False)


    x_values = avg_drop_df["Sweep"]

    ax.plot(x_values, avg_drop_df["Pulse 1 to 2 delta V"],
            label="Pulse 2 - Pulse 1")
    ax.plot(x_values, avg_drop_df["Pulse 2 to 3 delta V"],
            label="Pulse 3 - Pulse 2")
    ax.plot(x_values, avg_drop_df["Pulse 3 to 4 delta V"],
            label="Pulse 4 - Pulse 3")

    ax.set(xlabel="Sweep Number", ylabel="fEPSP ∆V (V)")
    fig.legend()
    fig.savefig("Plots/fEPSP ∆V vs Sweep Number")
    fig.show()
    plt.close(fig)

def main():
    initialize_directories()
    data_excel = pd.read_excel(data_path, sheet_name=None)

    # load all data
    recordings = []
    for sheet_name, sheet_df in data_excel.items():
        recordings.append(Recording(sheet_df))

    generate_pulse_fv_amplitude_vs_sweep(recordings)
    generate_avg_epsp_vs_sweep(recordings)
    generate_avg_epsp_vs_pulse_fv_normalized(recordings)
    generate_pulse_fv_time_vs_sweep(recordings)
    generate_epspToFV_ratios_pulse_comparisons(recordings)

    facilitation_df, facilitation_summary_df = summarize_facilitation_across_sweeps(
        recordings)

    generate_normalized_epsp_vs_sweep(recordings)
    generate_normalized_fv_vs_sweep(recordings)
    generate_avg_drop_fv_amplitude(recordings)
    generate_avg_drop_epsp_amplitude(recordings)


if __name__ == '__main__':
    main()
