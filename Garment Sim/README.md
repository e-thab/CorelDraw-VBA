# Garment Simulation

Automatically places artwork on garment photos of all sizes, styles, and colors to simulate printed product.

- File browser to change destination.
- Auto-filling options based on auto-detected garment type.
- Aesthetically pleasing representation of file path.
- Dynamic progress bar + informative progress tallies.

Consists of a form and module.

# How to Implement

- Place Bar.jpg and BarCancel.jpg somewhere fun, update path references.
- Remove all absolute file path references and replace with desired defaults.
- Configure export settings in GarmentSim(), GarmentSimCZ(), and ExportGarmCZ()

Similar story to Export and Publish, much of this was hard-coded from day 1 since it never needed to be any other way for internal use.
Customization will not be simple.

A streamlined, easy to customize version is on the docket.