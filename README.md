MOF–Bead LCA Model
This repository contains the life cycle assessment (LCA) calculation framework used to quantify the climate change impacts of producing bio-derived polymer beads and UiO-66-NH₂ functionalised composite beads for aqueous copper removal. The model supports the analysis presented in the associated journal paper and supplementary information.
Purpose
The tool evaluates how adding UiO-66-NH₂ functionality to a structured chitosan/PDChNF bead changes manufacturing energy demand, material inputs, solvent use and recovery, and climate impact (GWP100), including performance-normalised impact per gram of Cu removed. It is designed for transparent, parameter-driven screening at laboratory process scale with physically consistent scaling of electricity use.
Product systems
Ref-Bead: PDChNF–chitosan structured bead without MOF. U@Bead-2step-aUiO: Same support with two-step UiO-66-NH₂ growth.
System boundary
The foreground system is gate-to-gate, starting with the consumption of purchased inputs (chemicals and electricity) and ending when 1 kg of dry bead leaves the final drying step. Upstream burdens of purchased chemicals and grid electricity are included using cradle-to-gate background datasets.
Functional units
FU1: Production of 1 kg dry bead. FU2: Removal of 1 g Cu²⁺ from water. Adsorption capacities used are 77 g Cu kg⁻¹ bead for Ref-Bead and 116 g Cu kg⁻¹ bead for U@Bead-2step-aUiO.
Key modelling features
Explicit unit-operation electricity inventory (mixing, microfluidisation, centrifugation, crosslinking, drying, MOF growth steps). Laboratory electricity and utilisation-scaled electricity scenarios. Solvent recovery for ethanol and formic acid with recovery energy included. Separation of impacts into electricity, polymer support, and MOF growth inputs.
Intended use
This model is intended for comparing process routes, identifying environmental hotspots, testing sensitivity to scaling and recovery assumptions, and supporting research publications.
Limitations
This is not a full industrial plant design, not a cradle-to-grave assessment, and not a techno-economic model.
Citation
If you use this model, code, or derived results in academic work, please cite the associated publication and this repository.
Licence
This project is released under the BSD 3-Clause License.
