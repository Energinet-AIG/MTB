from TopFunction import execute


# Decide whether or not to plot the figures
plot = True
# Decide which results are plotted, RMS or EMT or both
plot_RMS = False
plot_EMT = True

# Decide whether or not to generate a report
generate_report = False


execute(
        plot_flag = plot,
        RMS_flag = plot_RMS,
        EMT_flag = plot_EMT,
        report_flag = generate_report,
        )
