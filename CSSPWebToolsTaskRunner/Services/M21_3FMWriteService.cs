using CSSPWebToolsTaskRunner.Services.Resources;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class M21_3FMService_Old
    {
        // --------------------
        // m21_3fm write section
        // --------------------

        private void Write_M21_3FM_Document(StreamWriter sw, M21_3FMService m21_3fmInput, FileInfo fi)
        {
            StringBuilder sb = new StringBuilder();

            try
            {
                M21_3FMService.TopFileInfo myTopFileInfo = this.topfileinfo;
                sw.WriteLine("// Created     : " + String.Format("{0:yyyy-MM-dd H:m:s}", myTopFileInfo.Created));
                sw.WriteLine("// DLL id      : " + myTopFileInfo.DLLid.ToString());
                sw.WriteLine("// PFS version : " + myTopFileInfo.PFSversion.ToString());
                sw.WriteLine("");
                M21_3FMService.FemEngineHD myFemEnginHD = this.femEngineHD;
                sw.WriteLine("[FemEngineHD]");
                M21_3FMService.FemEngineHD.DOMAIN myDomain = myFemEnginHD.domain;
                sw.WriteLine("   [DOMAIN]");
                sw.WriteLine("      Touched = " + myDomain.Touched.ToString());
                sw.WriteLine("      discretization = " + myDomain.discretization.ToString());
                sw.WriteLine("      number_of_dimensions = " + myDomain.number_of_dimensions.ToString());
                sw.WriteLine("      number_of_meshes = " + myDomain.number_of_meshes.ToString());
                sw.WriteLine("      file_name = " + myDomain.file_name.ToString());
                sw.WriteLine("      type_of_reordering = " + myDomain.type_of_reordering.ToString());
                sw.WriteLine("      number_of_domains = " + myDomain.number_of_domains.ToString());
                sw.WriteLine("      coordinate_type = " + myDomain.coordinate_type.ToString());
                sw.WriteLine("      minimum_depth = " + myDomain.minimum_depth.ToString());
                sw.WriteLine("      datum_depth = " + myDomain.datum_depth.ToString());
                sw.WriteLine("      vertical_mesh_type_overall = " + myDomain.vertical_mesh_type_overall.ToString());
                sw.WriteLine("      number_of_layers = " + myDomain.number_of_layers.ToString());
                sw.WriteLine("      z_sigma = " + myDomain.z_sigma.ToString());
                sw.WriteLine("      vertical_mesh_type = " + myDomain.vertical_mesh_type.ToString());
                sw.WriteLine(expandList("      layer_thickness = ", myDomain.layer_thickness));
                sw.WriteLine("      sigma_c = " + myDomain.sigma_c.ToString());
                sw.WriteLine("      theta = " + myDomain.theta.ToString());
                sw.WriteLine("      b = " + myDomain.b.ToString());
                sw.WriteLine("      number_of_layers_zlevel = " + myDomain.number_of_layers_zlevel.ToString());
                sw.WriteLine("      vertical_mesh_type_zlevel = " + myDomain.vertical_mesh_type_zlevel.ToString());
                sw.WriteLine("      constant_layer_thickness_zlevel = " + myDomain.constant_layer_thickness_zlevel.ToString());
                sw.WriteLine(expandList("      variable_layer_thickness_zlevel = ", myDomain.variable_layer_thickness_zlevel));
                sw.WriteLine("      type_of_bathymetry_adjustment = " + myDomain.type_of_bathymetry_adjustment.ToString());
                sw.WriteLine("      minimum_layer_thickness_zlevel = " + myDomain.minimum_layer_thickness_zlevel.ToString());
                sw.WriteLine("      type_of_mesh = " + myDomain.type_of_mesh.ToString());
                sw.WriteLine("      type_of_gauss = " + myDomain.type_of_gauss.ToString());
                M21_3FMService.FemEngineHD.DOMAIN.BOUNDARY_NAMES myBoundary_Names = myDomain.boundary_names;
                sw.WriteLine("      [BOUNDARY_NAMES]");
                sw.WriteLine("         Touched = " + myBoundary_Names.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myBoundary_Names.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.DOMAIN.BOUNDARY_NAMES.BOUNDARY_CODE> boundary_code in myBoundary_Names.boundary_code)
                {
                    sb.AppendLine("         [" + boundary_code.Key.ToString() + "]");
                    sb.AppendLine("            Touched = " + boundary_code.Value.Touched.ToString());
                    sb.AppendLine("            name = " + boundary_code.Value.Name.ToString());
                    sb.AppendLine("         EndSect  // " + boundary_code.Key.ToString());
                    sb.AppendLine("");
                }
                sw.WriteLine(sb.ToString().TrimEnd());
                sb.Clear();
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // BOUNDARY_NAMES");
                myBoundary_Names = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.DOMAIN.GIS_BACKGROUND myGis_Background = myDomain.gis_background;
                sw.WriteLine("      [GIS_BACKGROUND]");
                sw.WriteLine("         Touched = " + myGis_Background.Touched.ToString());
                sw.WriteLine("         file_name = " + myGis_Background.file_Name.ToString());
                sw.WriteLine("      EndSect  // GIS_BACKGROUND");
                myGis_Background = null;
                sw.WriteLine("");
                sw.WriteLine("   EndSect  // DOMAIN");
                myDomain = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TIME myTime_FEHD = myFemEnginHD.time;
                sw.WriteLine("   [TIME]");
                sw.WriteLine("      Touched = " + myTime_FEHD.Touched.ToString());
                sw.WriteLine("      start_time = " + String.Format("{0:yyyy, M, d, H, m, s}", myTime_FEHD.start_time));
                sw.WriteLine("      time_step_interval = " + myTime_FEHD.time_step_interval.ToString());
                sw.WriteLine("      number_of_time_steps = " + myTime_FEHD.number_of_time_steps.ToString());
                sw.WriteLine("   EndSect  // TIME");
                myTime_FEHD = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.MODULE_SELECTION myModule_Selection = myFemEnginHD.module_selection;
                sw.WriteLine("   [MODULE_SELECTION]");
                sw.WriteLine("      Touched = " + myModule_Selection.Touched.ToString());
                sw.WriteLine("      mode_of_hydrodynamic_module = " + myModule_Selection.mode_of_hydrodynamic_module.ToString());
                sw.WriteLine("      mode_of_spectral_wave_module = " + myModule_Selection.mode_of_spectral_wave_module.ToString());
                sw.WriteLine("      mode_of_transport_module = " + myModule_Selection.mode_of_transport_module.ToString());
                sw.WriteLine("      mode_of_mud_transport_module = " + myModule_Selection.mode_of_mud_transport_module.ToString());
                sw.WriteLine("      mode_of_eco_lab_module = " + myModule_Selection.mode_of_eco_lab_module.ToString());
                sw.WriteLine("      mode_of_sand_transport_module = " + myModule_Selection.mode_of_sand_transport_module.ToString());
                sw.WriteLine("      mode_of_particle_tracking_module = " + myModule_Selection.mode_of_particle_tracking_module.ToString());
                sw.WriteLine("      mode_of_oil_spill_module = " + myModule_Selection.mode_of_oil_spill_module.ToString());
                sw.WriteLine("   EndSect  // MODULE_SELECTION");
                myModule_Selection = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE myHydrodynamic_Module = myFemEnginHD.hydrodynamic_module;
                sw.WriteLine("   [HYDRODYNAMIC_MODULE]");
                sw.WriteLine("      mode = " + myHydrodynamic_Module.mode.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EQUATION myEquation = myHydrodynamic_Module.equation;
                sw.WriteLine("      [EQUATION]");
                sw.WriteLine("         formulation = " + myEquation.formulation.ToString());
                sw.WriteLine("      EndSect  // EQUATION");
                myEquation = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TIME myTime_Hm = myHydrodynamic_Module.time;
                sw.WriteLine("      [TIME]");
                sw.WriteLine("         start_time_step = " + myTime_Hm.start_time_step.ToString());
                sw.WriteLine("         time_step_factor = " + myTime_Hm.time_step_factor.ToString());
                sw.WriteLine("      EndSect  // TIME");
                myTime_Hm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SPACE mySpace_Hd = myHydrodynamic_Module.space;
                sw.WriteLine("      [SPACE]");
                sw.WriteLine("         number_of_2D_mesh_geometry = " + mySpace_Hd.number_of_2D_mesh_geometry.ToString());
                sw.WriteLine("         number_of_2D_mesh_velocity = " + mySpace_Hd.number_of_2D_mesh_velocity.ToString());
                sw.WriteLine("         number_of_2D_mesh_elevation = " + mySpace_Hd.number_of_2D_mesh_elevation.ToString());
                sw.WriteLine("         number_of_3D_mesh_geometry = " + mySpace_Hd.number_of_3D_mesh_geometry.ToString());
                sw.WriteLine("         number_of_3D_mesh_velocity = " + mySpace_Hd.number_of_3D_mesh_velocity.ToString());
                sw.WriteLine("         number_of_3D_mesh_pressure = " + mySpace_Hd.number_of_3D_mesh_pressure.ToString());
                sw.WriteLine("      EndSect  // SPACE");
                mySpace_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOLUTION_TECHNIQUE mySolution_Technique_Hd = myHydrodynamic_Module.solution_technique;
                sw.WriteLine("      [SOLUTION_TECHNIQUE]");
                sw.WriteLine("         Touched = " + mySolution_Technique_Hd.Touched.ToString());
                sw.WriteLine("         scheme_of_time_integration = " + ((int)mySolution_Technique_Hd.scheme_of_time_integration).ToString());
                sw.WriteLine("         scheme_of_space_discretization_horizontal = " + ((int)mySolution_Technique_Hd.scheme_of_space_discretization_horizontal).ToString());
                sw.WriteLine("         scheme_of_space_discretization_vertical = " + ((int)mySolution_Technique_Hd.scheme_of_space_discretization_vertical).ToString());
                sw.WriteLine("         method_of_space_discretization_horizontal = " + mySolution_Technique_Hd.method_of_space_discretization_horizontal.ToString());
                sw.WriteLine("         CFL_critical_HD = " + mySolution_Technique_Hd.CFL_critical_HD.ToString());
                sw.WriteLine("         dt_min_HD = " + mySolution_Technique_Hd.dt_min_HD.ToString());
                sw.WriteLine("         dt_max_HD = " + mySolution_Technique_Hd.dt_max_HD.ToString());
                sw.WriteLine("         CFL_critical_AD = " + mySolution_Technique_Hd.CFL_critical_AD.ToString());
                sw.WriteLine("         dt_min_AD = " + mySolution_Technique_Hd.dt_min_AD.ToString());
                sw.WriteLine("         dt_max_AD = " + mySolution_Technique_Hd.dt_max_AD.ToString());
                sw.WriteLine("         error_level = " + mySolution_Technique_Hd.error_level.ToString());
                sw.WriteLine("         maximum_number_of_errors = " + mySolution_Technique_Hd.maximum_number_of_errors.ToString());
                sw.WriteLine("         tau = " + mySolution_Technique_Hd.tau.ToString());
                sw.WriteLine("         theta = " + mySolution_Technique_Hd.theta.ToString());
                sw.WriteLine("         convection_bound = " + mySolution_Technique_Hd.convection_bound.ToString());
                sw.WriteLine("         artificial_diffusion = " + mySolution_Technique_Hd.artificial_diffusion.ToString());
                sw.WriteLine("         [DIFFERENTIAL_ALGEBRAIC_EQUATION_SYSTEM]");
                sw.WriteLine("         EndSect  // DIFFERENTIAL_ALGEBRAIC_EQUATION_SYSTEM");
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOLUTION_TECHNIQUE.LINEAR_EQUATION_SYSTEM myLinear_Equation_System = mySolution_Technique_Hd.linear_equation_system;
                sw.WriteLine("         [LINEAR_EQUATION_SYSTEM]");
                sw.WriteLine("            method = " + myLinear_Equation_System.method.ToString());
                sw.WriteLine("            tolerance = " + myLinear_Equation_System.tolerance.ToString());
                sw.WriteLine("            max_iterations = " + myLinear_Equation_System.max_iterations.ToString());
                sw.WriteLine("            type_of_print = " + myLinear_Equation_System.type_of_print.ToString());
                sw.WriteLine("         EndSect  // LINEAR_EQUATION_SYSTEM");
                myLinear_Equation_System = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // SOLUTION_TECHNIQUE");
                mySolution_Technique_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.FLOOD_AND_DRY myFlood_And_Dry = myHydrodynamic_Module.flood_and_dry;
                sw.WriteLine("      [FLOOD_AND_DRY]");
                sw.WriteLine("         Touched = " + myFlood_And_Dry.Touched.ToString());
                sw.WriteLine("         type = " + ((int)myFlood_And_Dry.type).ToString());
                sw.WriteLine("         drying_depth = " + myFlood_And_Dry.drying_depth.ToString());
                sw.WriteLine("         flooding_depth = " + myFlood_And_Dry.flooding_depth.ToString());
                sw.WriteLine("         mass_depth = " + myFlood_And_Dry.mass_depth.ToString());
                sw.WriteLine("         depth = " + myFlood_And_Dry.depth.ToString());
                sw.WriteLine("         max_depth = " + myFlood_And_Dry.max_depth.ToString());
                sw.WriteLine("         width = " + myFlood_And_Dry.width.ToString());
                sw.WriteLine("         smoothing_parameter = " + myFlood_And_Dry.smoothing_parameter.ToString());
                sw.WriteLine("         friction_coefficient = " + myFlood_And_Dry.friction_coefficient.ToString());
                sw.WriteLine("      EndSect  // FLOOD_AND_DRY");
                myFlood_And_Dry = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DEPTH myDepth = myHydrodynamic_Module.depth;
                sw.WriteLine("      [DEPTH]");
                sw.WriteLine("         Touched = " + myDepth.Touched.ToString());
                sw.WriteLine("      EndSect  // DEPTH");
                myDepth = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY myDensity = myHydrodynamic_Module.density;
                sw.WriteLine("      [DENSITY]");
                sw.WriteLine("         Touched = " + myDensity.Touched.ToString());
                sw.WriteLine("         type = " + ((int)myDensity.type).ToString());
                sw.WriteLine("         temperature_reference = " + myDensity.temperature_reference.ToString());
                sw.WriteLine("         salinity_reference = " + myDensity.salinity_reference.ToString());
                sw.WriteLine("      EndSect  // DENSITY");
                myDensity = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EDDY_VISCOSITY myEddy_Viscosity = myHydrodynamic_Module.eddy_viscosity;
                sw.WriteLine("      [EDDY_VISCOSITY]");
                sw.WriteLine("         Touched = " + myEddy_Viscosity.Touched.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EDDY_VISCOSITY.HORIZONTAL_EDDY_VISCOSITY myEddy_Viscosity_Horizontal = myEddy_Viscosity.horizontal_eddy_viscosity;
                sw.WriteLine("         [HORIZONTAL_EDDY_VISCOSITY]");
                sw.WriteLine("            Touched = " + myEddy_Viscosity_Horizontal.Touched.ToString());
                sw.WriteLine("            type = " + ((int)myEddy_Viscosity_Horizontal.type).ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EDDY_VISCOSITY.HORIZONTAL_EDDY_VISCOSITY.CONSTANT_EDDY_FORMULATION myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation = myEddy_Viscosity_Horizontal.constant_eddy_formulation;
                sw.WriteLine("            [CONSTANT_EDDY_FORMULATION]");
                sw.WriteLine("               Touched = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.Touched.ToString());
                sw.WriteLine("               type = " + ((int)myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.type).ToString());
                sw.WriteLine("               format = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.format.ToString());
                sw.WriteLine("               constant_value = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.constant_value.ToString());
                sw.WriteLine("               file_name = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.file_name.ToString());
                sw.WriteLine("               item_number = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.item_number.ToString());
                sw.WriteLine("               item_name = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // CONSTANT_EDDY_FORMULATION");
                myEddy_Viscosity_Horizontal_Constant_Eddy_Formulation = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EDDY_VISCOSITY.HORIZONTAL_EDDY_VISCOSITY.SMAGORINSKY_FORMULATION myEddy_Viscosity_Horizontal_Smagorinsky_Formulation = myEddy_Viscosity_Horizontal.smagorinsky_formulation;
                sw.WriteLine("            [SMAGORINSKY_FORMULATION]");
                sw.WriteLine("               Touched = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.Touched.ToString());
                sw.WriteLine("               type = " + ((int)myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.type).ToString());
                sw.WriteLine("               format = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.format.ToString());
                sw.WriteLine("               constant_value = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.constant_value.ToString());
                sw.WriteLine("               file_name = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.file_name.ToString());
                sw.WriteLine("               item_number = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.item_number.ToString());
                sw.WriteLine("               item_name = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.type_of_time_interpolation.ToString());
                sw.WriteLine("               minimum_eddy_viscosity = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.minimum_eddy_viscosity.ToString());
                sw.WriteLine("               maximum_eddy_viscosity = " + myEddy_Viscosity_Horizontal_Smagorinsky_Formulation.maximum_eddy_viscosity.ToString());
                sw.WriteLine("            EndSect  // SMAGORINSKY_FORMULATION");
                myEddy_Viscosity_Horizontal_Smagorinsky_Formulation = null;
                sw.WriteLine("");
                sw.WriteLine("         EndSect  // HORIZONTAL_EDDY_VISCOSITY");
                myEddy_Viscosity_Horizontal = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EDDY_VISCOSITY.VERTICAL_EDDY_VISCOSITY myEddy_Viscosity_Vertical = myEddy_Viscosity.vertical_eddy_viscosity;
                sw.WriteLine("         [VERTICAL_EDDY_VISCOSITY]");
                sw.WriteLine("            Touched = " + myEddy_Viscosity_Vertical.Touched.ToString());
                sw.WriteLine("            type = " + myEddy_Viscosity_Vertical.type.ToString());
                sw.WriteLine("            [CONSTANT_EDDY_FORMULATION]");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EDDY_VISCOSITY.VERTICAL_EDDY_VISCOSITY.CONSTANT_EDDY_FORMULATION myEddy_Viscosity_Vertical_Constant_Eddy_Formulation = myEddy_Viscosity_Vertical.constant_eddy_formulation;
                sw.WriteLine("               Touched = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.Touched.ToString());
                sw.WriteLine("               type = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.type.ToString());
                sw.WriteLine("               format = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.format.ToString());
                sw.WriteLine("               constant_value = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.constant_value.ToString());
                sw.WriteLine("               file_name = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.file_name.ToString());
                sw.WriteLine("               item_number = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.item_number.ToString());
                sw.WriteLine("               item_name = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.type_of_time_interpolation.ToString());
                sw.WriteLine("               Ri_damping = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.Ri_damping.ToString());
                sw.WriteLine("               Ri_a = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.Ri_a.ToString());
                sw.WriteLine("               Ri_b = " + myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.Ri_b.ToString());
                sw.WriteLine(expandList("               mixing_length_constants = ", myEddy_Viscosity_Vertical_Constant_Eddy_Formulation.mixing_length_constants));
                sw.WriteLine("            EndSect  // CONSTANT_EDDY_FORMULATION");
                myEddy_Viscosity_Vertical_Constant_Eddy_Formulation = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EDDY_VISCOSITY.VERTICAL_EDDY_VISCOSITY.LOG_LAW_FORMULATION myEddy_Viscosity_Vertical_Log_Law_Formulation = myEddy_Viscosity_Vertical.log_law_formulation;
                sw.WriteLine("            [LOG_LAW_FORMULATION]");
                sw.WriteLine("               Touched = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.Touched.ToString());
                sw.WriteLine("               type_of_Top_layer = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.type_of_Top_layer.ToString());
                sw.WriteLine("               thickness_of_Top_layer = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.thickness_of_Top_layer.ToString());
                sw.WriteLine("               fraction_of_depth = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.fraction_of_depth.ToString());
                sw.WriteLine("               fraction_of_Top_layer = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.fraction_of_Top_layer.ToString());
                sw.WriteLine("               type_of_Bottom_layer = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.type_of_Bottom_layer.ToString());
                sw.WriteLine("               thickness_of_Bottom_layer = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.thickness_of_Bottom_layer.ToString());
                sw.WriteLine("               fraction_of_depth = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.fraction_of_depth2.ToString());
                sw.WriteLine("               fraction_of_Bottom_layer = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.fraction_of_Bottom_layer.ToString());
                sw.WriteLine("               minimum_eddy_viscosity = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.minimum_eddy_viscosity.ToString());
                sw.WriteLine("               maximum_eddy_viscosity = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.maximum_eddy_viscosity.ToString());
                sw.WriteLine("               Ri_damping = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.Ri_damping.ToString());
                sw.WriteLine("               Ri_a = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.Ri_a.ToString());
                sw.WriteLine("               Ri_b = " + myEddy_Viscosity_Vertical_Log_Law_Formulation.Ri_b.ToString());
                sw.WriteLine(expandList("               mixing_length_constants = ", myEddy_Viscosity_Vertical_Log_Law_Formulation.mixing_length_constants));
                sw.WriteLine("            EndSect  // LOG_LAW_FORMULATION");
                myEddy_Viscosity_Vertical_Log_Law_Formulation = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.EDDY_VISCOSITY.VERTICAL_EDDY_VISCOSITY.K_EPSILON_FORMULATION myEddy_Viscosity_Vertical_K_Epsilon_Formulation = myEddy_Viscosity_Vertical.k_epsilon_formulation;
                sw.WriteLine("            [K_EPSILON_FORMULATION]");
                sw.WriteLine("               Touched = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.Touched.ToString());
                sw.WriteLine("               type_of_Top_layer = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.type_of_Top_layer.ToString());
                sw.WriteLine("               thickness_of_Top_layer = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.thickness_of_Top_layer.ToString());
                sw.WriteLine("               fraction_of_depth = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.fraction_of_depth.ToString());
                sw.WriteLine("               fraction_of_Top_layer = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.fraction_of_Top_layer.ToString());
                sw.WriteLine("               type_of_Bottom_layer = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.type_of_Bottom_layer.ToString());
                sw.WriteLine("               thickness_of_Bottom_layer = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.thickness_of_Bottom_layer.ToString());
                sw.WriteLine("               fraction_of_depth = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.fraction_of_depth2.ToString());
                sw.WriteLine("               fraction_of_Bottom_layer = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.fraction_of_Bottom_layer.ToString());
                sw.WriteLine("               minimum_eddy_viscosity = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.minimum_eddy_viscosity.ToString());
                sw.WriteLine("               maximum_eddy_viscosity = " + myEddy_Viscosity_Vertical_K_Epsilon_Formulation.maximum_eddy_viscosity.ToString());
                sw.WriteLine("            EndSect  // K_EPSILON_FORMULATION");
                myEddy_Viscosity_Vertical_K_Epsilon_Formulation = null;
                sw.WriteLine("");
                sw.WriteLine("         EndSect  // VERTICAL_EDDY_VISCOSITY");
                myEddy_Viscosity_Vertical = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // EDDY_VISCOSITY");
                myEddy_Viscosity = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE myBed_Resistance_Hd = myHydrodynamic_Module.bed_resistance;
                sw.WriteLine("      [BED_RESISTANCE]");
                sw.WriteLine("         Touched = " + myBed_Resistance_Hd.Touched.ToString());
                sw.WriteLine("         type = " + ((int)myBed_Resistance_Hd.type).ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.DRAG_COEFFICIENT myDragCoefficient = myBed_Resistance_Hd.drag_coefficient;
                sw.WriteLine("         [DRAG_COEFFICIENT]");
                sw.WriteLine("            Touched = " + myDragCoefficient.Touched.ToString());
                sw.WriteLine("            type = " + myDragCoefficient.type.ToString());
                sw.WriteLine("            format = " + myDragCoefficient.format.ToString());
                sw.WriteLine("            constant_value = " + myDragCoefficient.constant_value.ToString());
                sw.WriteLine("            file_name = " + myDragCoefficient.file_name.ToString());
                sw.WriteLine("            item_number = " + myDragCoefficient.item_number.ToString());
                sw.WriteLine("            item_name = " + myDragCoefficient.item_name.ToString());
                sw.WriteLine("            type_of_soft_start = " + myDragCoefficient.type_of_soft_start.ToString());
                sw.WriteLine("            soft_time_interval = " + myDragCoefficient.soft_time_interval.ToString());
                sw.WriteLine("            reference_value = " + myDragCoefficient.reference_value.ToString());
                sw.WriteLine("            type_of_time_interpolation = " + myDragCoefficient.type_of_time_interpolation.ToString());
                sw.WriteLine("         EndSect  // DRAG_COEFFICIENT");
                myDragCoefficient = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.CHEZY_NUMBER myChezy_Number = myBed_Resistance_Hd.chezy_number;
                sw.WriteLine("         [CHEZY_NUMBER]");
                sw.WriteLine("            Touched = " + myChezy_Number.Touched.ToString());
                sw.WriteLine("            type = " + myChezy_Number.type.ToString());
                sw.WriteLine("            format = " + myChezy_Number.format.ToString());
                sw.WriteLine("            constant_value = " + myChezy_Number.constant_value.ToString());
                sw.WriteLine("            file_name = " + myChezy_Number.file_name.ToString());
                sw.WriteLine("            item_number = " + myChezy_Number.item_number.ToString());
                sw.WriteLine("            item_name = " + myChezy_Number.item_name.ToString());
                sw.WriteLine("            type_of_soft_start = " + myChezy_Number.type_of_soft_start.ToString());
                sw.WriteLine("            soft_time_interval = " + myChezy_Number.soft_time_interval.ToString());
                sw.WriteLine("            reference_value = " + myChezy_Number.reference_value.ToString());
                sw.WriteLine("            type_of_time_interpolation = " + myChezy_Number.type_of_time_interpolation.ToString());
                sw.WriteLine("         EndSect  // CHEZY_NUMBER");
                myChezy_Number = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.MANNING_NUMBER myManning_Number = myBed_Resistance_Hd.manning_number;
                sw.WriteLine("         [MANNING_NUMBER]");
                sw.WriteLine("            Touched = " + myManning_Number.Touched.ToString());
                sw.WriteLine("            type = " + myManning_Number.type.ToString());
                sw.WriteLine("            format = " + myManning_Number.format.ToString());
                sw.WriteLine("            constant_value = " + myManning_Number.constant_value.ToString());
                sw.WriteLine("            file_name = " + myManning_Number.file_name.ToString());
                sw.WriteLine("            item_number = " + myManning_Number.item_number.ToString());
                sw.WriteLine("            item_name = " + myManning_Number.item_name.ToString());
                sw.WriteLine("            type_of_soft_start = " + myManning_Number.type_of_soft_start.ToString());
                sw.WriteLine("            soft_time_interval = " + myManning_Number.soft_time_interval.ToString());
                sw.WriteLine("            reference_value = " + myManning_Number.reference_value.ToString());
                sw.WriteLine("            type_of_time_interpolation = " + myManning_Number.type_of_time_interpolation.ToString());
                sw.WriteLine("            type_of_Bottom_layer = " + myManning_Number.type_of_Bottom_layer.ToString());
                sw.WriteLine("            thickness_of_Bottom_layer = " + myManning_Number.thickness_of_Bottom_layer.ToString());
                sw.WriteLine("            fraction_of_depth = " + myManning_Number.fraction_of_depth.ToString());
                sw.WriteLine("            fraction_of_Bottom_layer = " + myManning_Number.fraction_of_Bottom_layer.ToString());
                sw.WriteLine("         EndSect  // MANNING_NUMBER");
                myManning_Number = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.ROUGHNESS myRoughness = myBed_Resistance_Hd.roughness;
                sw.WriteLine("         [ROUGHNESS]");
                sw.WriteLine("            Touched = " + myRoughness.Touched.ToString());
                sw.WriteLine("            type = " + myRoughness.type.ToString());
                sw.WriteLine("            format = " + myRoughness.format.ToString());
                sw.WriteLine("            constant_value = " + myRoughness.constant_value.ToString());
                sw.WriteLine("            file_name = " + myRoughness.file_name.ToString());
                sw.WriteLine("            item_number = " + myRoughness.item_number.ToString());
                sw.WriteLine("            item_name = " + myRoughness.item_name.ToString());
                sw.WriteLine("            type_of_soft_start = " + myRoughness.type_of_soft_start.ToString());
                sw.WriteLine("            soft_time_interval = " + myRoughness.soft_time_interval.ToString());
                sw.WriteLine("            reference_value = " + myRoughness.reference_value.ToString());
                sw.WriteLine("            type_of_time_interpolation = " + myRoughness.type_of_time_interpolation.ToString());
                sw.WriteLine("            type_of_Bottom_layer = " + myRoughness.type_of_Bottom_layer.ToString());
                sw.WriteLine("            thickness_of_Bottom_layer = " + myRoughness.thickness_of_Bottom_layer.ToString());
                sw.WriteLine("            fraction_of_depth = " + myRoughness.fraction_of_depth.ToString());
                sw.WriteLine("            fraction_of_Bottom_layer = " + myRoughness.fraction_of_Bottom_layer.ToString());
                sw.WriteLine("         EndSect  // ROUGHNESS");
                myRoughness = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // BED_RESISTANCE");
                myBed_Resistance_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.CORIOLIS myCoriolis = myHydrodynamic_Module.coriolis;
                sw.WriteLine("      [CORIOLIS]");
                sw.WriteLine("         Touched = " + myCoriolis.Touched.ToString());
                sw.WriteLine("         type = " + ((int)myCoriolis.type).ToString());
                sw.WriteLine("         latitude = " + myCoriolis.latitude.ToString());
                sw.WriteLine("      EndSect  // CORIOLIS");
                myCoriolis = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.WIND_FORCING myWind_Forcing = myHydrodynamic_Module.wind_forcing;
                sw.WriteLine("      [WIND_FORCING]");
                sw.WriteLine("         type = " + ((int)myWind_Forcing.type).ToString());
                sw.WriteLine("         format = " + myWind_Forcing.format.ToString());
                sw.WriteLine("         constant_speed = " + myWind_Forcing.constant_speed.ToString());
                sw.WriteLine("         constant_direction = " + myWind_Forcing.constant_direction.ToString());
                sw.WriteLine("         file_name = " + myWind_Forcing.file_name.ToString());
                sw.WriteLine("         neutral_pressure = " + myWind_Forcing.neutral_pressure.ToString());
                sw.WriteLine("         type_of_soft_start = " + myWind_Forcing.type_of_soft_start.ToString());
                sw.WriteLine("         soft_time_interval = " + myWind_Forcing.soft_time_interval.ToString());
                sw.WriteLine("         [WIND_FRICTION]");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.WIND_FORCING.WIND_FRICTION myWind_Friction = myWind_Forcing.wind_friction;
                sw.WriteLine("            Touched = " + myWind_Friction.Touched.ToString());
                sw.WriteLine("            type = " + ((int)myWind_Friction.type).ToString());
                sw.WriteLine("            constant_friction = " + myWind_Friction.constant_friction.ToString());
                sw.WriteLine("            linear_friction_low = " + myWind_Friction.linear_friction_low.ToString());
                sw.WriteLine("            linear_friction_high = " + myWind_Friction.linear_friction_high.ToString());
                sw.WriteLine("            linear_speed_low = " + myWind_Friction.linear_speed_low.ToString());
                sw.WriteLine("            linear_speed_high = " + myWind_Friction.linear_speed_high.ToString());
                sw.WriteLine("         EndSect  // WIND_FRICTION");
                myWind_Friction = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // WIND_FORCING");
                myWind_Friction = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.ICE myIce = myHydrodynamic_Module.ice;
                sw.WriteLine("      [ICE]");
                sw.WriteLine("         Touched = " + myIce.Touched.ToString());
                sw.WriteLine("         type = " + ((int)myIce.type).ToString());
                sw.WriteLine("         format = " + myIce.format.ToString());
                sw.WriteLine("         c_cut_off = " + myIce.c_cut_off.ToString());
                sw.WriteLine("         file_name = " + myIce.file_name.ToString());
                sw.WriteLine("         item_number_concentration = " + myIce.item_number_concentration.ToString());
                sw.WriteLine("         item_number_thickness = " + myIce.item_number_thickness.ToString());
                sw.WriteLine("         item_name_concentration = " + myIce.item_name_concentration.ToString());
                sw.WriteLine("         item_name_thickness = " + myIce.item_name_thickness.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.ICE.ROUGHNESS myIce_Roughness = myIce.roughness;
                sw.WriteLine("         [ROUGHNESS]");
                sw.WriteLine("            Touched = " + myIce_Roughness.Touched.ToString());
                sw.WriteLine("            type = " + ((int)myIce_Roughness.type).ToString());
                sw.WriteLine("            format = " + myIce_Roughness.format.ToString());
                sw.WriteLine("            constant_value = " + myIce_Roughness.constant_value.ToString());
                sw.WriteLine("            file_name = " + myIce_Roughness.file_name.ToString());
                sw.WriteLine("            item_number = " + myIce_Roughness.item_number.ToString());
                sw.WriteLine("            item_name = " + myIce_Roughness.item_name.ToString());
                sw.WriteLine("            type_of_soft_start = " + myIce_Roughness.type_of_soft_start.ToString());
                sw.WriteLine("            soft_time_interval = " + myIce_Roughness.soft_time_interval.ToString());
                sw.WriteLine("            reference_value = " + myIce_Roughness.reference_value.ToString());
                sw.WriteLine("            type_of_time_interpolation = " + myIce_Roughness.type_of_time_interpolation.ToString());
                sw.WriteLine("         EndSect  // ROUGHNESS");
                myIce_Roughness = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // ICE");
                myIce = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TIDAL_POTENTIAL myTidal_Potential = myHydrodynamic_Module.tidal_potential;
                sw.WriteLine("      [TIDAL_POTENTIAL]");
                sw.WriteLine("         Touched = " + myTidal_Potential.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myTidal_Potential.MzSEPfsListItemCount.ToString());
                sw.WriteLine("         type = " + ((int)myTidal_Potential.type).ToString());
                sw.WriteLine("         format = " + myTidal_Potential.format.ToString());
                sw.WriteLine("         constituent_file_name = " + myTidal_Potential.constituent_file_name.ToString());
                sw.WriteLine("         number_of_constituents = " + myTidal_Potential.number_of_constituents.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TIDAL_POTENTIAL.CONSTITUENT> myConstituent in myTidal_Potential.constituents)
                {
                    sw.WriteLine("         [" + myConstituent.Key.ToString() + "]");
                    sw.WriteLine("            Touched = " + myConstituent.Value.Touched.ToString());
                    sw.WriteLine("            name = " + myConstituent.Value.name.ToString());
                    sw.WriteLine("            species = " + myConstituent.Value.species.ToString());
                    sw.WriteLine("            constituent = " + myConstituent.Value.constituent.ToString());
                    sw.WriteLine("            amplitude = " + myConstituent.Value.amplitude.ToString());
                    sw.WriteLine("            earthtide = " + myConstituent.Value.earthtide.ToString());
                    sw.WriteLine("            period_scaling = " + myConstituent.Value.period_scaling.ToString());
                    sw.WriteLine("            period = " + myConstituent.Value.period.ToString());
                    sw.WriteLine("            nodal_number_1 = " + myConstituent.Value.nodal_number_1.ToString());
                    sw.WriteLine("            nodal_number_2 = " + myConstituent.Value.nodal_number_2.ToString());
                    sw.WriteLine("            nodal_number_3 = " + myConstituent.Value.nodal_number_3.ToString());
                    sw.WriteLine(expandList("            arguments = ", myConstituent.Value.arguments));
                    sw.WriteLine("            phase = " + myConstituent.Value.phase.ToString());
                    sw.WriteLine("         EndSect  // " + myConstituent.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // TIDAL_POTENTIAL");
                myTidal_Potential = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.PRECIPITATION_EVAPORATION myPrecipitation_Evaporation_Hd = myHydrodynamic_Module.precipitation_evaporation;
                sw.WriteLine("      [PRECIPITATION_EVAPORATION]");
                sw.WriteLine("         Touched = " + myPrecipitation_Evaporation_Hd.Touched.ToString());
                sw.WriteLine("         type_of_precipitation = " + ((int)myPrecipitation_Evaporation_Hd.type_of_precipitation).ToString());
                sw.WriteLine("         type_of_evaporation = " + ((int)myPrecipitation_Evaporation_Hd.type_of_evaporation).ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.PRECIPITATION_EVAPORATION.PRECIPITATION myPrecipitation = myPrecipitation_Evaporation_Hd.precipitation;
                sw.WriteLine("         [PRECIPITATION]");
                sw.WriteLine("            Touched = " + myPrecipitation.Touched.ToString());
                sw.WriteLine("            type = " + myPrecipitation.type.ToString());
                sw.WriteLine("            format = " + myPrecipitation.format.ToString());
                sw.WriteLine("            constant_value = " + myPrecipitation.constant_value.ToString());
                sw.WriteLine("            file_name = " + myPrecipitation.file_name.ToString());
                sw.WriteLine("            item_number = " + myPrecipitation.item_number.ToString());
                sw.WriteLine("            item_name = " + myPrecipitation.item_name.ToString());
                sw.WriteLine("            type_of_soft_start = " + myPrecipitation.type_of_soft_start.ToString());
                sw.WriteLine("            soft_time_interval = " + myPrecipitation.soft_time_interval.ToString());
                sw.WriteLine("            reference_value = " + myPrecipitation.reference_value.ToString());
                sw.WriteLine("            type_of_time_interpolation = " + myPrecipitation.type_of_time_interpolation.ToString());
                sw.WriteLine("         EndSect  // PRECIPITATION");
                myPrecipitation = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.PRECIPITATION_EVAPORATION.EVAPORATION myEvaporation = myPrecipitation_Evaporation_Hd.evaporation;
                sw.WriteLine("         [EVAPORATION]");
                sw.WriteLine("            Touched = " + myEvaporation.Touched.ToString());
                sw.WriteLine("            type = " + myEvaporation.type.ToString());
                sw.WriteLine("            format = " + myEvaporation.format.ToString());
                sw.WriteLine("            constant_value = " + myEvaporation.constant_value.ToString());
                sw.WriteLine("            file_name = " + myEvaporation.file_name.ToString());
                sw.WriteLine("            item_number = " + myEvaporation.item_number.ToString());
                sw.WriteLine("            item_name = " + myEvaporation.item_name.ToString());
                sw.WriteLine("            type_of_soft_start = " + myEvaporation.type_of_soft_start.ToString());
                sw.WriteLine("            soft_time_interval = " + myEvaporation.soft_time_interval.ToString());
                sw.WriteLine("            reference_value = " + myEvaporation.reference_value.ToString());
                sw.WriteLine("            type_of_time_interpolation = " + myEvaporation.type_of_time_interpolation.ToString());
                sw.WriteLine("         EndSect  // EVAPORATION");
                myPrecipitation_Evaporation_Hd = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // PRECIPITATION_EVAPORATION");
                myPrecipitation_Evaporation_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.RADIATION_STRESS myRadiation_Stress = myHydrodynamic_Module.radiation_stress;
                sw.WriteLine("      [RADIATION_STRESS]");
                sw.WriteLine("         Touched = " + myRadiation_Stress.Touched.ToString());
                sw.WriteLine("         type = " + ((int)myRadiation_Stress.type).ToString());
                sw.WriteLine("         format = " + myRadiation_Stress.format.ToString());
                sw.WriteLine("         file_name = " + myRadiation_Stress.file_name.ToString());
                sw.WriteLine("         soft_time_interval = " + myRadiation_Stress.soft_time_interval.ToString());
                sw.WriteLine("      EndSect  // RADIATION_STRESS");
                myRadiation_Stress = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES mySources_Hd = myHydrodynamic_Module.sources;
                sw.WriteLine("      [SOURCES]");
                sw.WriteLine("         Touched = " + mySources_Hd.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + mySources_Hd.MzSEPfsListItemCount.ToString());
                sw.WriteLine("         number_of_sources = " + mySources_Hd.number_of_sources.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE> mySource_Hd in mySources_Hd.source)
                {
                    sw.WriteLine("         [" + mySource_Hd.Key.ToString() + "]");
                    sw.WriteLine("            Name = " + mySource_Hd.Value.Name.ToString());
                    sw.WriteLine("            include = " + mySource_Hd.Value.include.ToString());
                    sw.WriteLine("            interpolation_type = " + mySource_Hd.Value.interpolation_type.ToString());
                    sw.WriteLine("            coordinate_type = " + mySource_Hd.Value.coordinate_type.ToString());
                    sw.WriteLine("            zone = " + mySource_Hd.Value.zone.ToString());
                    sw.WriteLine("            coordinates = " + mySource_Hd.Value.coordinates.x.ToString() + ", " + mySource_Hd.Value.coordinates.y.ToString() + ", " + mySource_Hd.Value.coordinates.z.ToString());
                    sw.WriteLine("            layer = " + mySource_Hd.Value.layer.ToString());
                    sw.WriteLine("            distribution_type = " + mySource_Hd.Value.distribution_type.ToString());
                    sw.WriteLine("            connected_source = " + mySource_Hd.Value.connected_source.ToString());
                    sw.WriteLine("            type = " + ((int)mySource_Hd.Value.type).ToString());
                    sw.WriteLine("            format = " + mySource_Hd.Value.format.ToString());
                    sw.WriteLine("            file_name = " + mySource_Hd.Value.file_name.ToString());
                    sw.WriteLine("            constant_value = " + mySource_Hd.Value.constant_value.ToString());
                    sw.WriteLine("            item_number = " + mySource_Hd.Value.item_number.ToString());
                    sw.WriteLine("            item_name = " + mySource_Hd.Value.item_name.ToString());
                    sw.WriteLine("            type_of_soft_start = " + mySource_Hd.Value.type_of_soft_start.ToString());
                    sw.WriteLine("            soft_time_interval = " + mySource_Hd.Value.soft_time_interval.ToString());
                    sw.WriteLine("            reference_value = " + mySource_Hd.Value.reference_value.ToString());
                    sw.WriteLine("            type_of_time_interpolation = " + mySource_Hd.Value.type_of_time_interpolation.ToString());
                    sw.WriteLine("         EndSect  // " + mySource_Hd.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // SOURCES");
                mySources_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE myStructure_Module_Hd = myHydrodynamic_Module.structure_module;
                sw.WriteLine("      [STRUCTURE_MODULE]");
                sw.WriteLine(expandList("         Structure_Version = ", myStructure_Module_Hd.Structure_Version));
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.CROSSSECTIONS myCrosssections_Hd = myStructure_Module_Hd.crosssections;
                sw.WriteLine("         [CROSSSECTIONS]");
                sw.WriteLine("            CrossSectionDataBridge = " + myCrosssections_Hd.CrossSectionDataBridge.ToString());
                sw.WriteLine("            CrossSectionFile = " + myCrosssections_Hd.CrossSectionFile.ToString());
                sw.WriteLine("         EndSect  // CROSSSECTIONS");
                myCrosssections_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.WEIR myWeir_Hd = myStructure_Module_Hd.weir;
                sw.WriteLine("         [WEIR]");
                foreach (M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.WEIR.WEIR_DATA myWeir_Weir_Data_Hd in myWeir_Hd.weir_data)
                {
                    sw.WriteLine("            [weir_data]");
                    sw.WriteLine(expandList("               Location = ", myWeir_Weir_Data_Hd.Location));
                    sw.WriteLine("               delhs = " + myWeir_Weir_Data_Hd.delhs.ToString());
                    sw.WriteLine("               coordinate_type = " + myWeir_Weir_Data_Hd.coordinate_type.ToString());
                    sw.WriteLine("               number_of_points = " + myWeir_Weir_Data_Hd.number_of_points.ToString());
                    sw.WriteLine(expandList("               x = ", myWeir_Weir_Data_Hd.x));
                    sw.WriteLine(expandList("               y = ", myWeir_Weir_Data_Hd.y));
                    sw.WriteLine("               HorizOffset = " + myWeir_Weir_Data_Hd.HorizOffset.ToString());
                    sw.WriteLine("               Attributes = "
                        + ((int)myWeir_Weir_Data_Hd.attributes.type).ToString()
                        + ", " + ((int)myWeir_Weir_Data_Hd.attributes.valve).ToString());
                    sb.Clear();
                    sw.WriteLine(expandList("               HeadLossFactors = ", myWeir_Weir_Data_Hd.HeadLossFactors));
                    sw.WriteLine(expandList("               WeirFormulaParam = ", myWeir_Weir_Data_Hd.WeirFormulaParam));
                    sw.WriteLine(expandList("               WeirFormula2Param = ", myWeir_Weir_Data_Hd.WeirFormula2Param));
                    sw.WriteLine(expandList("               WeirFormula3Param = ", myWeir_Weir_Data_Hd.WeirFormula3Param));
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.WEIR.WEIR_DATA.GEOMETRY myWeir_Weir_Data_Geometry_Hd = myWeir_Weir_Data_Hd.Geometry;
                    sw.WriteLine("               [Geometry]");
                    sw.WriteLine(expandList("                  Attributes = ", myWeir_Weir_Data_Geometry_Hd.Attributes));
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.WEIR.WEIR_DATA.GEOMETRY.LEVEL_WIDTH myWeir_Weir_Data_Geometry_Level_Width_Hd = myWeir_Weir_Data_Geometry_Hd.Level_Width;
                    sw.WriteLine("                  [Level_Width]");
                    sw.WriteLine("                  EndSect  // Level_Width");
                    myWeir_Weir_Data_Geometry_Level_Width_Hd = null;
                    sw.WriteLine("");
                    sw.WriteLine("               EndSect  // Geometry");
                    myWeir_Weir_Data_Geometry_Hd = null;
                    sw.WriteLine("");
                    sw.WriteLine("            EndSect  // weir_data");
                    sw.WriteLine("");
                }
                sw.WriteLine("         EndSect  // WEIR");
                sw.WriteLine("");
                myWeir_Hd = null;
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.CULVERTS myCulverts_Hd = myStructure_Module_Hd.culverts;
                sw.WriteLine("         [CULVERTS]");
                foreach (M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.CULVERTS.CULVERT_DATA myCulverts_Hd_Culvert_Data_Hd in myCulverts_Hd.culvert_data)
                {
                    sw.WriteLine("            [culvert_data]");
                    sw.WriteLine(expandList("               Location = ", myCulverts_Hd_Culvert_Data_Hd.Location));
                    sw.WriteLine("               delhs = " + myCulverts_Hd_Culvert_Data_Hd.delhs.ToString());
                    sw.WriteLine("               coordinate_type = " + myCulverts_Hd_Culvert_Data_Hd.coordinate_type.ToString());
                    sw.WriteLine("               number_of_points = " + myCulverts_Hd_Culvert_Data_Hd.number_of_points.ToString());
                    sw.WriteLine(expandList("               x = ", myCulverts_Hd_Culvert_Data_Hd.x));
                    sw.WriteLine(expandList("               y = ", myCulverts_Hd_Culvert_Data_Hd.y));
                    sw.WriteLine("               HorizOffset = " + myCulverts_Hd_Culvert_Data_Hd.HorizOffset.ToString());
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.CULVERTS.CULVERT_DATA.ATTRIBUTES myCulverts_Hd_Culvert_Data_Hd_attributes =
                        myCulverts_Hd_Culvert_Data_Hd.attributes;
                    sw.WriteLine("               Attributes = " +
                        myCulverts_Hd_Culvert_Data_Hd_attributes.Upstream.ToString() + ", " +
                        myCulverts_Hd_Culvert_Data_Hd_attributes.Downstream.ToString() + ", " +
                        myCulverts_Hd_Culvert_Data_Hd_attributes.Length.ToString() + ", " +
                        myCulverts_Hd_Culvert_Data_Hd_attributes.Manning_n.ToString() + ", " +
                        myCulverts_Hd_Culvert_Data_Hd_attributes.NumberOfCulverts.ToString() + ", " +
                        ((int)myCulverts_Hd_Culvert_Data_Hd_attributes.valve_regulation).ToString() + ", " +
                        ((int)myCulverts_Hd_Culvert_Data_Hd_attributes.section_type).ToString());
                    myCulverts_Hd_Culvert_Data_Hd_attributes = null;
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURE_MODULE.CULVERTS.CULVERT_DATA.GEOMETRY myCulverts_Hd_Culvert_Data_Hd_Geometry_Hd = myCulverts_Hd_Culvert_Data_Hd.geometry;
                    sw.WriteLine("               [Geometry]");
                    sw.WriteLine("                  Type = " + ((int)myCulverts_Hd_Culvert_Data_Hd_Geometry_Hd.Type).ToString());
                    sw.WriteLine(expandList("                  Rectangular = ", myCulverts_Hd_Culvert_Data_Hd_Geometry_Hd.Rectangular));
                    sw.WriteLine("                  Cicular_Diameter = " + myCulverts_Hd_Culvert_Data_Hd_Geometry_Hd.Cicular_Diameter);
                    sw.WriteLine("                  [Irregular]");
                    sw.WriteLine("                  EndSect  // Irregular");
                    sw.WriteLine("");
                    sw.WriteLine("               EndSect  // Geometry");
                    myCulverts_Hd_Culvert_Data_Hd_Geometry_Hd = null;
                    sw.WriteLine("");
                    sw.WriteLine(expandList("               HeadLossFactors = ", myCulverts_Hd_Culvert_Data_Hd.HeadLossFactors));
                    sw.WriteLine("            EndSect  // culvert_data");
                    sw.WriteLine("");
                }
                sw.WriteLine("         EndSect  // CULVERTS");
                myCulverts_Hd = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // STRUCTURE_MODULE");
                myStructure_Module_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES myStructures_Hd = myHydrodynamic_Module.structures;
                sw.WriteLine("      [STRUCTURES]");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES.GATES myGates_Hd = myStructures_Hd.gates;
                sw.WriteLine("         [GATES]");
                sw.WriteLine("            Touched = " + myGates_Hd.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + myGates_Hd.MzSEPfsListItemCount.ToString());
                sw.WriteLine("            number_of_gates = " + myGates_Hd.number_of_gates.ToString());
                if (myGates_Hd.gate != null)
                {
                    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES.GATES.GATE> myGates_Hd_Gate_Hd in myGates_Hd.gate)
                    {
                        sw.WriteLine("            [" + myGates_Hd_Gate_Hd.Key.ToString() + "]");
                        sw.WriteLine("               Name = " + myGates_Hd_Gate_Hd.Value.Name.ToString());
                        sw.WriteLine("               include = " + myGates_Hd_Gate_Hd.Value.include.ToString());
                        sw.WriteLine("               input_format = " + myGates_Hd_Gate_Hd.Value.input_format.ToString());
                        sw.WriteLine("               coordinate_type = " + myGates_Hd_Gate_Hd.Value.coordinate_type.ToString());
                        sw.WriteLine("               number_of_points = " + myGates_Hd_Gate_Hd.Value.number_of_points.ToString());
                        sw.WriteLine(expandList("               x = ", myGates_Hd_Gate_Hd.Value.x));
                        sw.WriteLine(expandList("               y = ", myGates_Hd_Gate_Hd.Value.y));
                        foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES.GATES.GATE.POINT> myGates_Hd_Gate_Hd_Point_Hd in myGates_Hd_Gate_Hd.Value.point)
                        {
                            sw.WriteLine("               [" + myGates_Hd_Gate_Hd_Point_Hd.Key.ToString() + "]");
                            sw.WriteLine("                  x = " + myGates_Hd_Gate_Hd_Point_Hd.Value.x);
                            sw.WriteLine("                  y = " + myGates_Hd_Gate_Hd_Point_Hd.Value.y);
                            sw.WriteLine("               EndSect  // " + myGates_Hd_Gate_Hd_Point_Hd.Key.ToString());
                            sw.WriteLine("");
                        }
                        sw.WriteLine("               input_file_name = " + myGates_Hd_Gate_Hd.Value.input_file_name.ToString());
                        sw.WriteLine("               format = " + myGates_Hd_Gate_Hd.Value.format.ToString());
                        sw.WriteLine("               constant_value = " + myGates_Hd_Gate_Hd.Value.constant_value.ToString());
                        sw.WriteLine("               file_name = " + myGates_Hd_Gate_Hd.Value.file_name.ToString());
                        sw.WriteLine("               item_number = " + myGates_Hd_Gate_Hd.Value.item_number.ToString());
                        sw.WriteLine("               item_name = " + myGates_Hd_Gate_Hd.Value.item_name.ToString());
                        sw.WriteLine("               type_of_soft_start = " + myGates_Hd_Gate_Hd.Value.type_of_soft_start.ToString());
                        sw.WriteLine("               soft_time_interval = " + myGates_Hd_Gate_Hd.Value.soft_time_interval.ToString());
                        sw.WriteLine("               reference_value = " + myGates_Hd_Gate_Hd.Value.reference_value.ToString());
                        sw.WriteLine("               type_of_time_interpolation = " + myGates_Hd_Gate_Hd.Value.type_of_time_interpolation.ToString());
                        sw.WriteLine("            EndSect  // " + myGates_Hd_Gate_Hd.Key);
                        sw.WriteLine("");
                    }
                }
                sw.WriteLine("         EndSect  // GATES");
                myGates_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES.PIERS myPiers_Hd = myStructures_Hd.piers;
                sw.WriteLine("         [PIERS]");
                sw.WriteLine("            Touched = " + myPiers_Hd.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + myPiers_Hd.MzSEPfsListItemCount.ToString());
                sw.WriteLine("            format = " + myPiers_Hd.format.ToString());
                sw.WriteLine("            number_of_piers = " + myPiers_Hd.number_of_piers.ToString());
                if (myPiers_Hd.pier != null)
                {
                    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES.PIERS.PIER> myPiers_Hd_Pier_Hd in myPiers_Hd.pier)
                    {
                        sw.WriteLine("            [" + myPiers_Hd_Pier_Hd.Key.ToString() + "]");
                        sw.WriteLine("               Name = " + myPiers_Hd_Pier_Hd.Value.Name.ToString());
                        sw.WriteLine("               include = " + myPiers_Hd_Pier_Hd.Value.include.ToString());
                        sw.WriteLine("               coordinate_type = " + myPiers_Hd_Pier_Hd.Value.coordinate_type.ToString());
                        sw.WriteLine("               x = " + myPiers_Hd_Pier_Hd.Value.x.ToString());
                        sw.WriteLine("               y = " + myPiers_Hd_Pier_Hd.Value.y.ToString());
                        sw.WriteLine("               theta = " + myPiers_Hd_Pier_Hd.Value.theta.ToString());
                        sw.WriteLine("               lamda = " + myPiers_Hd_Pier_Hd.Value.lamda.ToString());
                        sw.WriteLine("               number_of_sections = " + myPiers_Hd_Pier_Hd.Value.number_of_sections.ToString());
                        sb = new StringBuilder();
                        sb.Append("               type = ");
                        for (int i = 0; i <= myPiers_Hd_Pier_Hd.Value.type.Count - 1; i++)
                        {
                            sb.Append(((int)myPiers_Hd_Pier_Hd.Value.type[i]).ToString().Trim() + ", ");
                        }
                        sb.Remove(sb.Length - 2, 1);
                        sw.WriteLine(sb.ToString().TrimEnd());
                        sw.WriteLine(expandList("               height = ", myPiers_Hd_Pier_Hd.Value.height));
                        sw.WriteLine(expandList("               length = ", myPiers_Hd_Pier_Hd.Value.length));
                        sw.WriteLine(expandList("               width = ", myPiers_Hd_Pier_Hd.Value.width));
                        sw.WriteLine(expandList("               radious = ", myPiers_Hd_Pier_Hd.Value.radious));
                        sw.WriteLine("            EndSect  // " + myPiers_Hd_Pier_Hd.Key.ToString());
                        sw.WriteLine("");
                    }
                }
                sw.WriteLine("         EndSect  // PIERS");
                myPiers_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES.TURBINES myTurbines_Hd = myStructures_Hd.turbines;
                sw.WriteLine("         [TURBINES]");
                sw.WriteLine("            Touched = " + myTurbines_Hd.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + myTurbines_Hd.MzSEPfsListItemCount.ToString());
                sw.WriteLine("            format = " + myTurbines_Hd.format.ToString());
                sw.WriteLine("            number_of_turbines = " + myTurbines_Hd.number_of_turbines.ToString());
                sw.WriteLine("            output_type = " + myTurbines_Hd.output_type.ToString());
                sw.WriteLine("            output_frequency = " + myTurbines_Hd.output_frequency.ToString());
                sw.WriteLine("            output_file_name = " + myTurbines_Hd.output_file_name.ToString());
                if (myTurbines_Hd.turbine != null)
                {
                    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES.TURBINES.TURBINE> myTurbines_Hd_Turbine_Hd in myTurbines_Hd.turbine)
                    {
                        sw.WriteLine("            [" + myTurbines_Hd_Turbine_Hd.Key.ToString() + "]");
                        sw.WriteLine("               Name = " + myTurbines_Hd_Turbine_Hd.Value.Name.ToString());
                        sw.WriteLine("               include = " + myTurbines_Hd_Turbine_Hd.Value.include.ToString());
                        sw.WriteLine("               coordinate_type = " + myTurbines_Hd_Turbine_Hd.Value.coordinate_type.ToString());
                        sw.WriteLine("               x = " + myTurbines_Hd_Turbine_Hd.Value.x.ToString());
                        sw.WriteLine("               y = " + myTurbines_Hd_Turbine_Hd.Value.y.ToString());
                        sw.WriteLine("               diameter = " + myTurbines_Hd_Turbine_Hd.Value.diameter.ToString());
                        sw.WriteLine("               centroid = " + myTurbines_Hd_Turbine_Hd.Value.centroid.ToString());
                        sw.WriteLine("               drag_coefficient = " + myTurbines_Hd_Turbine_Hd.Value.drag_coefficient.ToString());
                        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.STRUCTURES.TURBINES.TURBINE.CORRECTION_FACTOR myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd = myTurbines_Hd_Turbine_Hd.Value.correction_factor;
                        sw.WriteLine("               [CORRECTION_FACTOR]");
                        sw.WriteLine("                  Touched = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.Touched.ToString());
                        sw.WriteLine("                  format = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.format.ToString());
                        sw.WriteLine("                  constant_value = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.constant_value.ToString());
                        sw.WriteLine("                  file_name = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.file_name.ToString());
                        sw.WriteLine("                  item_number = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.item_number.ToString());
                        sw.WriteLine("                  item_name = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.item_name.ToString());
                        sw.WriteLine("                  type_of_soft_start = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.type_of_soft_start.ToString());
                        sw.WriteLine("                  soft_time_interval = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.soft_time_interval.ToString());
                        sw.WriteLine("                  reference_value = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.reference_value.ToString());
                        sw.WriteLine("                  type_of_time_interpolation = " + myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd.type_of_time_interpolation.ToString());
                        sw.WriteLine("               EndSect  // CORRECTION_FACTOR");
                        myTurbines_Hd_Turbine_Hd_Correction_Factor_Hd = null;
                        sw.WriteLine("");
                        sw.WriteLine("            EndSect  // " + myTurbines_Hd_Turbine_Hd.Key.ToString());
                        sw.WriteLine("");
                    }
                }
                sw.WriteLine("         EndSect  // TURBINES");
                myTurbines_Hd = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // STRUCTURES");
                myStructures_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.INITIAL_CONDITIONS myInitial_Conditions_Hd = myHydrodynamic_Module.initial_conditions;
                sw.WriteLine("      [INITIAL_CONDITIONS]");
                sw.WriteLine("         Touched = " + myInitial_Conditions_Hd.Touched.ToString());
                sw.WriteLine("         type = " + myInitial_Conditions_Hd.type.ToString());
                sw.WriteLine("         surface_elevation_constant = " + myInitial_Conditions_Hd.surface_elevation_constant.ToString());
                sw.WriteLine("         u_velocity_constant = " + myInitial_Conditions_Hd.u_velocity_constant.ToString());
                sw.WriteLine("         v_velocity_constant = " + myInitial_Conditions_Hd.v_velocity_constant.ToString());
                sw.WriteLine("         w_velocity_constant = " + myInitial_Conditions_Hd.w_velocity_constant.ToString());
                sw.WriteLine("      EndSect  // INITIAL_CONDITIONS");
                myInitial_Conditions_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BOUNDARY_CONDITIONS myBoundary_Conditions_Hd = myHydrodynamic_Module.boundary_conditions;

                sw.WriteLine("      [BOUNDARY_CONDITIONS]");
                sw.WriteLine("         Touched = " + myBoundary_Conditions_Hd.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myBoundary_Conditions_Hd.MzSEPfsListItemCount.ToString());
                sw.WriteLine("         internal_land_boundary_Type = " + myBoundary_Conditions_Hd.internal_land_boundary_Type.ToString());

                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BOUNDARY_CONDITIONS.CODE> myBoundary_Conditions_Hd_Code_Hd in myBoundary_Conditions_Hd.code)
                {
                    sw.WriteLine("         [" + myBoundary_Conditions_Hd_Code_Hd.Key.ToString() + "]");
                    if (myBoundary_Conditions_Hd_Code_Hd.Key.ToString().Trim().ToUpper() == "CODE_1")
                    {
                        sw.WriteLine("            type = " + ((int)myBoundary_Conditions_Hd_Code_Hd.Value.type).ToString());
                    }
                    else
                    {
                        sw.WriteLine("            identifier = " + myBoundary_Conditions_Hd_Code_Hd.Value.identifier.ToString());
                        sw.WriteLine("            type = " + ((int)myBoundary_Conditions_Hd_Code_Hd.Value.type).ToString());
                        sw.WriteLine("            type_interpolation_constrain = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_interpolation_constrain.ToString());
                        sw.WriteLine("            type_secondary = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_secondary.ToString());
                        sw.WriteLine("            type_of_vertical_profile = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_of_vertical_profile.ToString());
                        sw.WriteLine("            format = " + ((int)myBoundary_Conditions_Hd_Code_Hd.Value.format).ToString());
                        if (myBoundary_Conditions_Hd_Code_Hd.Value.constant_values != null)
                        {
                            sw.WriteLine(expandList("            constant_values = ", myBoundary_Conditions_Hd_Code_Hd.Value.constant_values));
                        }
                        else
                        {
                            sw.WriteLine("            constant_value = " + myBoundary_Conditions_Hd_Code_Hd.Value.constant_value.ToString());
                        }
                        sw.WriteLine("            file_name = " + myBoundary_Conditions_Hd_Code_Hd.Value.file_name.ToString());
                        if (myBoundary_Conditions_Hd_Code_Hd.Value.item_numbers != null)
                        {
                            sw.WriteLine(expandList("            item_numbers = ", myBoundary_Conditions_Hd_Code_Hd.Value.item_numbers));
                        }
                        else
                        {
                            sw.WriteLine("            item_number = " + myBoundary_Conditions_Hd_Code_Hd.Value.item_number.ToString());
                        }
                        if (myBoundary_Conditions_Hd_Code_Hd.Value.item_names != null)
                        {
                            sw.WriteLine(expandList("            item_names = ", myBoundary_Conditions_Hd_Code_Hd.Value.item_names));
                        }
                        else
                        {
                            sw.WriteLine("            item_name = " + myBoundary_Conditions_Hd_Code_Hd.Value.item_name.ToString());
                        }
                        sw.WriteLine("            type_of_soft_start = " + ((int)myBoundary_Conditions_Hd_Code_Hd.Value.type_of_soft_start).ToString());
                        sw.WriteLine("            soft_time_interval = " + myBoundary_Conditions_Hd_Code_Hd.Value.soft_time_interval.ToString());
                        if (myBoundary_Conditions_Hd_Code_Hd.Value.item_names != null)
                        {
                            sw.WriteLine(expandList("            reference_values = ", myBoundary_Conditions_Hd_Code_Hd.Value.reference_values));
                        }
                        else
                        {
                            sw.WriteLine("            reference_value = " + myBoundary_Conditions_Hd_Code_Hd.Value.reference_value.ToString());
                        }
                        sw.WriteLine("            type_of_time_interpolation = " + ((int)myBoundary_Conditions_Hd_Code_Hd.Value.type_of_time_interpolation).ToString());
                        sw.WriteLine("            type_of_space_interpolation = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_of_space_interpolation.ToString());
                        sw.WriteLine("            type_of_coriolis_correction = " + ((int)myBoundary_Conditions_Hd_Code_Hd.Value.type_of_coriolis_correction).ToString());
                        sw.WriteLine("            type_of_wind_correction = " + ((int)myBoundary_Conditions_Hd_Code_Hd.Value.type_of_wind_correction).ToString());
                        sw.WriteLine("            type_of_tilting = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_of_tilting.ToString());
                        sw.WriteLine("            type_of_tilting_point = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_of_tilting_point.ToString());
                        sw.WriteLine("            point_tilting = " + myBoundary_Conditions_Hd_Code_Hd.Value.point_tilting.ToString());
                        sw.WriteLine("            type_of_radiation_stress_correction = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_of_radiation_stress_correction.ToString());
                        sw.WriteLine("            type_of_pressure_correction = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_of_pressure_correction.ToString());
                        sw.WriteLine("            type_of_radiation_stress_correction = " + myBoundary_Conditions_Hd_Code_Hd.Value.type_of_radiation_stress_correction2.ToString());
                    }
                    sw.WriteLine("         EndSect  // " + myBoundary_Conditions_Hd_Code_Hd.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // BOUNDARY_CONDITIONS");
                myBoundary_Conditions_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE myTemperature_Salinity_Module = myHydrodynamic_Module.temperature_salinity_module;
                sw.WriteLine("      [TEMPERATURE_SALINITY_MODULE]");
                sw.WriteLine("         temperature_mode = " + myTemperature_Salinity_Module.temperature_mode.ToString());
                sw.WriteLine("         salinity_mode = " + myTemperature_Salinity_Module.salinity_mode.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.TIME myTime = myTemperature_Salinity_Module.time;
                sw.WriteLine("         [TIME]");
                sw.WriteLine("            start_time_step = " + myTime.start_time_step.ToString());
                sw.WriteLine("            time_step_factor = " + myTime.time_step_factor.ToString());
                sw.WriteLine("         EndSect  // TIME");
                myTime = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.SPACE mySpace_Tsm = myTemperature_Salinity_Module.space;
                sw.WriteLine("         [SPACE]");
                sw.WriteLine("            number_of_2D_mesh_geometry = " + mySpace_Tsm.number_of_2D_mesh_geometry.ToString());
                sw.WriteLine("            number_of_3D_mesh_geometry = " + mySpace_Tsm.number_of_3D_mesh_geometry.ToString());
                sw.WriteLine("         EndSect  // SPACE");
                mySpace_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.EQUATION myEquation_Tsm = myTemperature_Salinity_Module.equation_TSM;
                sw.WriteLine("         [EQUATION]");
                sw.WriteLine("            Touched = " + myEquation_Tsm.Touched.ToString());
                sw.WriteLine("            minimum_temperature = " + myEquation_Tsm.minimum_temperature.ToString());
                sw.WriteLine("            maximum_temperature = " + myEquation_Tsm.maximum_temperature.ToString());
                sw.WriteLine("            minimum_salinity = " + myEquation_Tsm.minimum_salinity.ToString());
                sw.WriteLine("            maximum_salinity = " + myEquation_Tsm.maximum_salinity.ToString());
                sw.WriteLine("         EndSect  // EQUATION");
                myEquation_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.SOLUTION_TECHNIQUE mySolution_Technique_Tsm = myTemperature_Salinity_Module.solution_technique;
                sw.WriteLine("         [SOLUTION_TECHNIQUE]");
                sw.WriteLine("            Touched = " + mySolution_Technique_Tsm.Touched.ToString());
                sw.WriteLine("            scheme_of_time_integration = " + ((int)mySolution_Technique_Tsm.scheme_of_time_integration).ToString());
                sw.WriteLine("            scheme_of_space_discretization_horizontal = " + mySolution_Technique_Tsm.scheme_of_space_discretization_horizontal.ToString());
                sw.WriteLine("            scheme_of_space_discretization_vertical = " + mySolution_Technique_Tsm.scheme_of_space_discretization_vertical.ToString());
                sw.WriteLine("            method_of_space_discretization_horizontal = " + mySolution_Technique_Tsm.method_of_space_discretization_horizontal.ToString());
                sw.WriteLine("         EndSect  // SOLUTION_TECHNIQUE");
                mySolution_Technique_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.DIFFUSION myDiffusion_Tsm = myTemperature_Salinity_Module.diffusion;
                sw.WriteLine("         [DIFFUSION]");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.DIFFUSION.HORIZONTAL_DIFFUSION myHorizontal_Diffusion_Tsm = myDiffusion_Tsm.horizontal_diffusion;
                sw.WriteLine("            [HORIZONTAL_DIFFUSION]");
                sw.WriteLine("               Touched = " + myHorizontal_Diffusion_Tsm.Touched.ToString());
                sw.WriteLine("               type = " + ((int)myHorizontal_Diffusion_Tsm.type).ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.DIFFUSION.HORIZONTAL_DIFFUSION.SCALED_EDDY_VISCOSITY myScaled_Eddy_Viscosity_Tsm = myHorizontal_Diffusion_Tsm.scaled_eddy_viscosity;
                sw.WriteLine("               [SCALED_EDDY_VISCOSITY]");
                sw.WriteLine("                  format = " + myScaled_Eddy_Viscosity_Tsm.format.ToString());
                sw.WriteLine("                  sigma = " + myScaled_Eddy_Viscosity_Tsm.sigma.ToString());
                sw.WriteLine("                  file_name = " + myScaled_Eddy_Viscosity_Tsm.file_name.ToString());
                sw.WriteLine("                  item_number = " + myScaled_Eddy_Viscosity_Tsm.item_number.ToString());
                sw.WriteLine("                  item_name = " + myScaled_Eddy_Viscosity_Tsm.item_name.ToString());
                sw.WriteLine("               EndSect  // SCALED_EDDY_VISCOSITY");
                myScaled_Eddy_Viscosity_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.DIFFUSION.HORIZONTAL_DIFFUSION.DIFFUSION_COEFFICIENT myHorizontal_Diffusion_Coefficient_Tsm = myHorizontal_Diffusion_Tsm.diffusion_coefficient;
                sw.WriteLine("               [DIFFUSION_COEFFICIENT]");
                sw.WriteLine("                  format = " + myHorizontal_Diffusion_Coefficient_Tsm.format.ToString());
                sw.WriteLine("                  constant_value = " + myHorizontal_Diffusion_Coefficient_Tsm.constant_value.ToString());
                sw.WriteLine("                  file_name = " + myHorizontal_Diffusion_Coefficient_Tsm.file_name.ToString());
                sw.WriteLine("                  item_number = " + myHorizontal_Diffusion_Coefficient_Tsm.item_number.ToString());
                sw.WriteLine("                  item_name = " + myHorizontal_Diffusion_Coefficient_Tsm.item_name.ToString());
                sw.WriteLine("               EndSect  // DIFFUSION_COEFFICIENT");
                myHorizontal_Diffusion_Coefficient_Tsm = null;
                sw.WriteLine("");
                sw.WriteLine("            EndSect  // HORIZONTAL_DIFFUSION");
                myHorizontal_Diffusion_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.DIFFUSION.VERTICAL_DIFFUSION myVertical_Diffusion_Tsm = myDiffusion_Tsm.vertical_diffusion;
                sw.WriteLine("            [VERTICAL_DIFFUSION]");
                sw.WriteLine("               Touched = " + myVertical_Diffusion_Tsm.Touched.ToString());
                sw.WriteLine("               type = " + myVertical_Diffusion_Tsm.type.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.DIFFUSION.VERTICAL_DIFFUSION.SCALED_EDDY_VISCOSITY myScaled_Eddy_ViscosityTsm = myVertical_Diffusion_Tsm.scaled_eddy_viscosity;
                sw.WriteLine("               [SCALED_EDDY_VISCOSITY]");
                sw.WriteLine("                  format = " + myScaled_Eddy_ViscosityTsm.format.ToString());
                sw.WriteLine("                  sigma = " + myScaled_Eddy_ViscosityTsm.sigma.ToString());
                sw.WriteLine("                  file_name = " + myScaled_Eddy_ViscosityTsm.file_name.ToString());
                sw.WriteLine("                  item_number = " + myScaled_Eddy_ViscosityTsm.item_number.ToString());
                sw.WriteLine("                  item_name = " + myScaled_Eddy_ViscosityTsm.item_name.ToString());
                sw.WriteLine("               EndSect  // SCALED_EDDY_VISCOSITY");
                myScaled_Eddy_ViscosityTsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.DIFFUSION.VERTICAL_DIFFUSION.DIFFUSION_COEFFICIENT myVertical_Diffusion_Coefficient_Tsm = myVertical_Diffusion_Tsm.diffusion_coefficient;
                sw.WriteLine("               [DIFFUSION_COEFFICIENT]");
                sw.WriteLine("                  format = " + myVertical_Diffusion_Coefficient_Tsm.format.ToString());
                sw.WriteLine("                  constant_value = " + myVertical_Diffusion_Coefficient_Tsm.constant_value.ToString());
                sw.WriteLine("                  file_name = " + myVertical_Diffusion_Coefficient_Tsm.file_name.ToString());
                sw.WriteLine("                  item_number = " + myVertical_Diffusion_Coefficient_Tsm.item_number.ToString());
                sw.WriteLine("                  item_name = " + myVertical_Diffusion_Coefficient_Tsm.item_name.ToString());
                sw.WriteLine("               EndSect  // DIFFUSION_COEFFICIENT");
                myVertical_Diffusion_Coefficient_Tsm = null;
                sw.WriteLine("");
                sw.WriteLine("            EndSect  // VERTICAL_DIFFUSION");
                myVertical_Diffusion_Tsm = null;
                sw.WriteLine("");
                sw.WriteLine("         EndSect  // DIFFUSION");
                myDiffusion_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.HEAT_EXCHANGE myHeat_Exchange_Tsm = myTemperature_Salinity_Module.heat_exchange;
                sw.WriteLine("         [HEAT_EXCHANGE]");
                sw.WriteLine("            Touched = " + myHeat_Exchange_Tsm.Touched.ToString());
                sw.WriteLine("            type = " + ((int)myHeat_Exchange_Tsm.type).ToString());
                sw.WriteLine("            Angstroms_law_A = " + myHeat_Exchange_Tsm.Angstroms_law_A.ToString());
                sw.WriteLine("            Angstroms_law_B = " + myHeat_Exchange_Tsm.Angstroms_law_B.ToString());
                sw.WriteLine("            Beers_law_beta = " + myHeat_Exchange_Tsm.Beers_law_beta.ToString());
                sw.WriteLine("            light_extinction = " + myHeat_Exchange_Tsm.light_extinction.ToString());
                sw.WriteLine("            displacement_hours = " + myHeat_Exchange_Tsm.displacement_hours.ToString());
                sw.WriteLine("            standard_meridian = " + myHeat_Exchange_Tsm.standard_meridian.ToString());
                sw.WriteLine("            Daltons_law_A = " + myHeat_Exchange_Tsm.Daltons_law_A.ToString());
                sw.WriteLine("            Daltons_law_B = " + myHeat_Exchange_Tsm.Daltons_law_B.ToString());
                sw.WriteLine("            sensible_heat_transfer_coefficient_heating = " + myHeat_Exchange_Tsm.sensible_heat_transfer_coefficient_heating.ToString());
                sw.WriteLine("            sensible_heat_transfer_coefficient_cooling = " + myHeat_Exchange_Tsm.sensible_heat_transfer_coefficient_cooling.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.HEAT_EXCHANGE.AIR_TEMPERATURE myAirTemperature_Tsm = myHeat_Exchange_Tsm.air_temperature;
                sw.WriteLine("            [air_temperature]");
                sw.WriteLine("               Touched = " + myAirTemperature_Tsm.Touched.ToString());
                sw.WriteLine("               type = " + myAirTemperature_Tsm.type.ToString());
                sw.WriteLine("               format = " + myAirTemperature_Tsm.format.ToString());
                sw.WriteLine("               constant_value = " + myAirTemperature_Tsm.constant_value.ToString());
                sw.WriteLine("               file_name = " + myAirTemperature_Tsm.file_name.ToString());
                sw.WriteLine("               item_number = " + myAirTemperature_Tsm.item_number.ToString());
                sw.WriteLine("               item_name = " + myAirTemperature_Tsm.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myAirTemperature_Tsm.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myAirTemperature_Tsm.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myAirTemperature_Tsm.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myAirTemperature_Tsm.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // air_temperature");
                myAirTemperature_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.HEAT_EXCHANGE.RELATIVE_HUMIDITY myRelative_Humidity_Tsm = myHeat_Exchange_Tsm.relative_humidity;
                sw.WriteLine("            [relative_humidity]");
                sw.WriteLine("               Touched = " + myRelative_Humidity_Tsm.Touched.ToString());
                sw.WriteLine("               type = " + myRelative_Humidity_Tsm.type.ToString());
                sw.WriteLine("               format = " + myRelative_Humidity_Tsm.format.ToString());
                sw.WriteLine("               constant_value = " + myRelative_Humidity_Tsm.constant_value.ToString());
                sw.WriteLine("               file_name = " + myRelative_Humidity_Tsm.file_name.ToString());
                sw.WriteLine("               item_number = " + myRelative_Humidity_Tsm.item_number.ToString());
                sw.WriteLine("               item_name = " + myRelative_Humidity_Tsm.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myRelative_Humidity_Tsm.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myRelative_Humidity_Tsm.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myRelative_Humidity_Tsm.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myRelative_Humidity_Tsm.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // relative_humidity");
                myRelative_Humidity_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.HEAT_EXCHANGE.CLEARNESS_COEFFICIENT myClearness_Coefficient_Tsm = myHeat_Exchange_Tsm.clearness_coefficient;
                sw.WriteLine("            [clearness_coefficient]");
                sw.WriteLine("               Touched = " + myClearness_Coefficient_Tsm.Touched.ToString());
                sw.WriteLine("               type = " + myClearness_Coefficient_Tsm.type.ToString());
                sw.WriteLine("               format = " + myClearness_Coefficient_Tsm.format.ToString());
                sw.WriteLine("               constant_value = " + myClearness_Coefficient_Tsm.constant_value.ToString());
                sw.WriteLine("               file_name = " + myClearness_Coefficient_Tsm.file_name.ToString());
                sw.WriteLine("               item_number = " + myClearness_Coefficient_Tsm.item_number.ToString());
                sw.WriteLine("               item_name = " + myClearness_Coefficient_Tsm.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myClearness_Coefficient_Tsm.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myClearness_Coefficient_Tsm.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myClearness_Coefficient_Tsm.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myClearness_Coefficient_Tsm.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // clearness_coefficient");
                myClearness_Coefficient_Tsm = null;
                sw.WriteLine("");
                sw.WriteLine("         EndSect  // HEAT_EXCHANGE");
                myHeat_Exchange_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.PRECIPITATION_EVAPORATION myPrecipitation_Evaporation_Tsm = myTemperature_Salinity_Module.precipitation_evaporation;
                sw.WriteLine("         [PRECIPITATION_EVAPORATION]");
                sw.WriteLine("            Touched = " + myPrecipitation_Evaporation_Tsm.Touched.ToString());
                sw.WriteLine("            type_of_precipitation = " + ((int)myPrecipitation_Evaporation_Tsm.type_of_precipitation).ToString());
                sw.WriteLine("            type_of_evaporation = " + myPrecipitation_Evaporation_Tsm.type_of_evaporation.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.PRECIPITATION_EVAPORATION.PRECIPITATION myPrecipitation_Tsm = myPrecipitation_Evaporation_Tsm.precipitation;
                sw.WriteLine("            [PRECIPITATION]");
                sw.WriteLine("               Touched = " + myPrecipitation_Tsm.Touched.ToString());
                sw.WriteLine("               type = " + myPrecipitation_Tsm.type.ToString());
                sw.WriteLine("               format = " + myPrecipitation_Tsm.format.ToString());
                sw.WriteLine("               constant_value = " + myPrecipitation_Tsm.constant_value.ToString());
                sw.WriteLine("               file_name = " + myPrecipitation_Tsm.file_name.ToString());
                sw.WriteLine("               item_number = " + myPrecipitation_Tsm.item_number.ToString());
                sw.WriteLine("               item_name = " + myPrecipitation_Tsm.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myPrecipitation_Tsm.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myPrecipitation_Tsm.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myPrecipitation_Tsm.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myPrecipitation_Tsm.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // PRECIPITATION");
                myPrecipitation_Tsm = null;
                sw.WriteLine("");
                sw.WriteLine("            [EVAPORATION]");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.PRECIPITATION_EVAPORATION.EVAPORATION myEvaporation_Tsm = myPrecipitation_Evaporation_Tsm.evaporation;
                sw.WriteLine("               Touched = " + myEvaporation_Tsm.Touched.ToString());
                sw.WriteLine("               type = " + myEvaporation_Tsm.type.ToString());
                sw.WriteLine("               format = " + myEvaporation_Tsm.format.ToString());
                sw.WriteLine("               constant_value = " + myEvaporation_Tsm.constant_value.ToString());
                sw.WriteLine("               file_name = " + myEvaporation_Tsm.file_name.ToString());
                sw.WriteLine("               item_number = " + myEvaporation_Tsm.item_number.ToString());
                sw.WriteLine("               item_name = " + myEvaporation_Tsm.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myEvaporation_Tsm.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myEvaporation_Tsm.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myEvaporation_Tsm.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myEvaporation_Tsm.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // EVAPORATION");
                myEvaporation_Tsm = null;
                sw.WriteLine("");
                sw.WriteLine("         EndSect  // PRECIPITATION_EVAPORATION");
                myPrecipitation_Evaporation_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.SOURCES mySources_Tsm = myTemperature_Salinity_Module.sources;
                sw.WriteLine("         [SOURCES]");
                sw.WriteLine("            Touched = " + mySources_Tsm.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + mySources_Tsm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.SOURCES.SOURCE> mySource_Tsm in mySources_Tsm.source)
                {
                    sw.WriteLine("            [" + mySource_Tsm.Key.ToString() + "]");
                    sw.WriteLine("               name = " + mySource_Tsm.Value.name.ToString());
                    sw.WriteLine("               type_of_temperature = " + mySource_Tsm.Value.type_of_temperature.ToString());
                    sw.WriteLine("               type_of_salinity = " + mySource_Tsm.Value.type_of_salinity.ToString());
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.SOURCES.SOURCE.TEMPERATURE myTemperature_Tsm = mySource_Tsm.Value.temperature;
                    sw.WriteLine("               [TEMPERATURE]");
                    sw.WriteLine("                  Touched = " + myTemperature_Tsm.Touched.ToString());
                    sw.WriteLine("                  type = " + myTemperature_Tsm.type.ToString());
                    sw.WriteLine("                  format = " + myTemperature_Tsm.format.ToString());
                    sw.WriteLine("                  constant_value = " + myTemperature_Tsm.constant_value.ToString());
                    sw.WriteLine("                  file_name = " + myTemperature_Tsm.file_name.ToString());
                    sw.WriteLine("                  item_number = " + myTemperature_Tsm.item_number.ToString());
                    sw.WriteLine("                  item_name = " + myTemperature_Tsm.item_name.ToString());
                    sw.WriteLine("                  type_of_soft_start = " + myTemperature_Tsm.type_of_soft_start.ToString());
                    sw.WriteLine("                  soft_time_interval = " + myTemperature_Tsm.soft_time_interval.ToString());
                    sw.WriteLine("                  reference_value = " + myTemperature_Tsm.reference_value.ToString());
                    sw.WriteLine("                  type_of_time_interpolation = " + myTemperature_Tsm.type_of_time_interpolation.ToString());
                    sw.WriteLine("               EndSect  // TEMPERATURE");
                    myTemperature_Tsm = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.SOURCES.SOURCE.SALINITY mySalinity_Tsm = mySource_Tsm.Value.salinity;
                    sw.WriteLine("               [SALINITY]");
                    sw.WriteLine("                  Touched = " + mySalinity_Tsm.Touched.ToString());
                    sw.WriteLine("                  type = " + mySalinity_Tsm.type.ToString());
                    sw.WriteLine("                  format = " + mySalinity_Tsm.format.ToString());
                    sw.WriteLine("                  constant_value = " + mySalinity_Tsm.constant_value.ToString());
                    sw.WriteLine("                  file_name = " + mySalinity_Tsm.file_name.ToString());
                    sw.WriteLine("                  item_number = " + mySalinity_Tsm.item_number.ToString());
                    sw.WriteLine("                  item_name = " + mySalinity_Tsm.item_name.ToString());
                    sw.WriteLine("                  type_of_soft_start = " + mySalinity_Tsm.type_of_soft_start.ToString());
                    sw.WriteLine("                  soft_time_interval = " + mySalinity_Tsm.soft_time_interval.ToString());
                    sw.WriteLine("                  reference_value = " + mySalinity_Tsm.reference_value.ToString());
                    sw.WriteLine("                  type_of_time_interpolation = " + mySalinity_Tsm.type_of_time_interpolation.ToString());
                    sw.WriteLine("               EndSect  // SALINITY");
                    mySalinity_Tsm = null;
                    sw.WriteLine("");
                    sw.WriteLine("            EndSect  // " + mySource_Tsm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("         EndSect  // SOURCES");
                mySources_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.INITIAL_CONDITIONS myInitial_Conditions_Tsm = myTemperature_Salinity_Module.initial_conditions;
                sw.WriteLine("         [INITIAL_CONDITIONS]");
                sw.WriteLine("            Touched = " + myInitial_Conditions_Tsm.Touched.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.INITIAL_CONDITIONS.TEMPERATURE_TSM myInitial_Conditions_Temperature_Tsm = myInitial_Conditions_Tsm.temperature;
                sw.WriteLine("            [TEMPERATURE]");
                sw.WriteLine("               Touched = " + myInitial_Conditions_Temperature_Tsm.Touched.ToString());
                sw.WriteLine("               type = " + myInitial_Conditions_Temperature_Tsm.type.ToString());
                sw.WriteLine("               format = " + myInitial_Conditions_Temperature_Tsm.format.ToString());
                sw.WriteLine("               constant_value = " + myInitial_Conditions_Temperature_Tsm.constant_value.ToString());
                sw.WriteLine("               file_name = " + myInitial_Conditions_Temperature_Tsm.file_name.ToString());
                sw.WriteLine("               item_number = " + myInitial_Conditions_Temperature_Tsm.item_number.ToString());
                sw.WriteLine("               item_name = " + myInitial_Conditions_Temperature_Tsm.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myInitial_Conditions_Temperature_Tsm.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myInitial_Conditions_Temperature_Tsm.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myInitial_Conditions_Temperature_Tsm.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myInitial_Conditions_Temperature_Tsm.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // TEMPERATURE");
                myInitial_Conditions_Temperature_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.INITIAL_CONDITIONS.SALINITY_TSM mySalinity_Tsn = myInitial_Conditions_Tsm.salinity;
                sw.WriteLine("            [SALINITY]");
                sw.WriteLine("               Touched = " + mySalinity_Tsn.Touched.ToString());
                sw.WriteLine("               type = " + mySalinity_Tsn.type.ToString());
                sw.WriteLine("               format = " + mySalinity_Tsn.format.ToString());
                sw.WriteLine("               constant_value = " + mySalinity_Tsn.constant_value.ToString());
                sw.WriteLine("               file_name = " + mySalinity_Tsn.file_name.ToString());
                sw.WriteLine("               item_number = " + mySalinity_Tsn.item_number.ToString());
                sw.WriteLine("               item_name = " + mySalinity_Tsn.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + mySalinity_Tsn.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + mySalinity_Tsn.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + mySalinity_Tsn.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + mySalinity_Tsn.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // SALINITY");
                mySalinity_Tsn = null;
                sw.WriteLine("");
                sw.WriteLine("         EndSect  // INITIAL_CONDITIONS");
                myInitial_Conditions_Tsm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.BOUNDARY_CONDITIONS myBoundary_Conditions_Tsm = myTemperature_Salinity_Module.boundary_conditions;
                sw.WriteLine("         [BOUNDARY_CONDITIONS]");
                sw.WriteLine("            Touched = " + myBoundary_Conditions_Tsm.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + myBoundary_Conditions_Tsm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.BOUNDARY_CONDITIONS.CODE> myCode_Bc in myBoundary_Conditions_Tsm.code)
                {
                    sw.WriteLine("            [" + myCode_Bc.Key.ToString() + "]");
                    if (myCode_Bc.Key.ToString().Trim().ToUpper() == "CODE_1")
                    {
                        sw.WriteLine("               [TEMPERATURE]");
                        sw.WriteLine("               EndSect  // TEMPERATURE");
                        sw.WriteLine("");
                        sw.WriteLine("               [SALINITY]");
                        sw.WriteLine("               EndSect  // SALINITY");
                        sw.WriteLine("");
                    }
                    else
                    {
                        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.BOUNDARY_CONDITIONS.CODE.TEMPERATURE myTemperature_Bc = myCode_Bc.Value.temperature;
                        sw.WriteLine("               [TEMPERATURE]");
                        sw.WriteLine("                  identifier = " + myTemperature_Bc.identifier.ToString());
                        sw.WriteLine("                  type = " + myTemperature_Bc.type.ToString());
                        sw.WriteLine("                  type_interpolation_constrain = " + myTemperature_Bc.type_interpolation_constrain.ToString());
                        sw.WriteLine("                  type_secondary = " + myTemperature_Bc.type_secondary.ToString());
                        sw.WriteLine("                  type_of_vertical_profile = " + myTemperature_Bc.type_of_vertical_profile.ToString());
                        sw.WriteLine("                  format = " + myTemperature_Bc.format.ToString());
                        sw.WriteLine("                  constant_value = " + myTemperature_Bc.constant_value.ToString());
                        sw.WriteLine("                  file_name = " + myTemperature_Bc.file_name.ToString());
                        sw.WriteLine("                  item_number = " + myTemperature_Bc.item_number.ToString());
                        sw.WriteLine("                  item_name = " + myTemperature_Bc.item_name.ToString());
                        sw.WriteLine("                  type_of_soft_start = " + myTemperature_Bc.type_of_soft_start.ToString());
                        sw.WriteLine("                  soft_time_interval = " + myTemperature_Bc.soft_time_interval.ToString());
                        sw.WriteLine("                  reference_value = " + myTemperature_Bc.reference_value.ToString());
                        sw.WriteLine("                  type_of_time_interpolation = " + myTemperature_Bc.type_of_time_interpolation.ToString());
                        sw.WriteLine("                  type_of_space_interpolation = " + myTemperature_Bc.type_of_space_interpolation.ToString());
                        sw.WriteLine("                  type_of_coriolis_correction = " + myTemperature_Bc.type_of_coriolis_correction.ToString());
                        sw.WriteLine("                  type_of_wind_correction = " + myTemperature_Bc.type_of_wind_correction.ToString());
                        sw.WriteLine("                  type_of_tilting = " + myTemperature_Bc.type_of_tilting.ToString());
                        sw.WriteLine("                  type_of_tilting_point = " + myTemperature_Bc.type_of_tilting_point.ToString());
                        sw.WriteLine("                  point_tilting = " + myTemperature_Bc.point_tilting.ToString());
                        sw.WriteLine("                  type_of_radiation_stress_correction = " + myTemperature_Bc.type_of_radiation_stress_correction.ToString());
                        sw.WriteLine("                  type_of_pressure_correction = " + myTemperature_Bc.type_of_pressure_correction.ToString());
                        sw.WriteLine("                  type_of_radiation_stress_correction = " + myTemperature_Bc.type_of_radiation_stress_correction2.ToString());
                        sw.WriteLine("               EndSect  // TEMPERATURE");
                        myTemperature_Bc = null;
                        sw.WriteLine("");
                        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TEMPERATURE_SALINITY_MODULE.BOUNDARY_CONDITIONS.CODE.SALINITY mySalinity_Bc = myCode_Bc.Value.salinity;
                        sw.WriteLine("               [SALINITY]");
                        sw.WriteLine("                  identifier = " + mySalinity_Bc.identifier.ToString());
                        sw.WriteLine("                  type = " + mySalinity_Bc.type.ToString());
                        sw.WriteLine("                  type_interpolation_constrain = " + mySalinity_Bc.type_interpolation_constrain.ToString());
                        sw.WriteLine("                  type_secondary = " + mySalinity_Bc.type_secondary.ToString());
                        sw.WriteLine("                  type_of_vertical_profile = " + mySalinity_Bc.type_of_vertical_profile.ToString());
                        sw.WriteLine("                  format = " + mySalinity_Bc.format.ToString());
                        sw.WriteLine("                  constant_value = " + mySalinity_Bc.constant_value.ToString());
                        sw.WriteLine("                  file_name = " + mySalinity_Bc.file_name.ToString());
                        sw.WriteLine("                  item_number = " + mySalinity_Bc.item_number.ToString());
                        sw.WriteLine("                  item_name = " + mySalinity_Bc.item_name.ToString());
                        sw.WriteLine("                  type_of_soft_start = " + mySalinity_Bc.type_of_soft_start.ToString());
                        sw.WriteLine("                  soft_time_interval = " + mySalinity_Bc.soft_time_interval.ToString());
                        sw.WriteLine("                  reference_value = " + mySalinity_Bc.reference_value.ToString());
                        sw.WriteLine("                  type_of_time_interpolation = " + mySalinity_Bc.type_of_time_interpolation.ToString());
                        sw.WriteLine("                  type_of_space_interpolation = " + mySalinity_Bc.type_of_space_interpolation.ToString());
                        sw.WriteLine("                  type_of_coriolis_correction = " + mySalinity_Bc.type_of_coriolis_correction.ToString());
                        sw.WriteLine("                  type_of_wind_correction = " + mySalinity_Bc.type_of_wind_correction.ToString());
                        sw.WriteLine("                  type_of_tilting = " + mySalinity_Bc.type_of_tilting.ToString());
                        sw.WriteLine("                  type_of_tilting_point = " + mySalinity_Bc.type_of_tilting_point.ToString());
                        sw.WriteLine("                  point_tilting = " + mySalinity_Bc.point_tilting.ToString());
                        sw.WriteLine("                  type_of_radiation_stress_correction = " + mySalinity_Bc.type_of_radiation_stress_correction.ToString());
                        sw.WriteLine("                  type_of_pressure_correction = " + mySalinity_Bc.type_of_pressure_correction.ToString());
                        sw.WriteLine("                  type_of_radiation_stress_correction = " + mySalinity_Bc.type_of_radiation_stress_correction2.ToString());
                        mySalinity_Bc = null;
                        sw.WriteLine("               EndSect  // SALINITY");
                        sw.WriteLine("");
                    }
                    sw.WriteLine("            EndSect  // " + myCode_Bc.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("         EndSect  // BOUNDARY_CONDITIONS");
                myBoundary_Conditions_Tsm = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // TEMPERATURE_SALINITY_MODULE");
                myTemperature_Salinity_Module = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE myTurbulence_Module = myHydrodynamic_Module.turbulence_module;
                sw.WriteLine("      [TURBULENCE_MODULE]");
                sw.WriteLine("         mode = " + myTurbulence_Module.mode.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.TIME myTurbulence_Module_Time = myTurbulence_Module.time;
                sw.WriteLine("         [TIME]");
                sw.WriteLine("            start_time_step = " + myTurbulence_Module_Time.start_time_step.ToString());
                sw.WriteLine("            time_step_factor = " + myTurbulence_Module_Time.time_step_factor.ToString());
                sw.WriteLine("         EndSect  // TIME");
                myTurbulence_Module_Time = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.SPACE mySpace_Tm = myTurbulence_Module.space;
                sw.WriteLine("         [SPACE]");
                sw.WriteLine("            number_of_2D_mesh_geometry = " + mySpace_Tm.number_of_2D_mesh_geometry.ToString());
                sw.WriteLine("            number_of_3D_mesh_geometry = " + mySpace_Tm.number_of_3D_mesh_geometry.ToString());
                sw.WriteLine("         EndSect  // SPACE");
                mySpace_Tm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.EQUATION myEquation_Tm = myTurbulence_Module.equation;
                sw.WriteLine("         [EQUATION]");
                sw.WriteLine("            Touched = " + myEquation_Tm.Touched.ToString());
                sw.WriteLine("            c1e = " + myEquation_Tm.c1e.ToString());
                sw.WriteLine("            c2e = " + myEquation_Tm.c2e.ToString());
                sw.WriteLine("            c3e = " + myEquation_Tm.c3e.ToString());
                sw.WriteLine("            prandtl_number = " + myEquation_Tm.prandtl_number.ToString());
                sw.WriteLine("            cmy = " + myEquation_Tm.cmy.ToString());
                sw.WriteLine("            minimum_kinetic_energy = " + myEquation_Tm.minimum_kinetic_energy.ToString());
                sw.WriteLine("            maximum_kinetic_energy = " + myEquation_Tm.maximum_kinetic_energy.ToString());
                sw.WriteLine("            minimum_dissipation_of_kinetic_energy = " + myEquation_Tm.minimum_dissipation_of_kinetic_energy.ToString());
                sw.WriteLine("            maximum_dissipation_of_kinetic_energy = " + myEquation_Tm.maximum_dissipation_of_kinetic_energy.ToString());
                sw.WriteLine("            surface_dissipation_parameter = " + myEquation_Tm.surface_dissipation_parameter.ToString());
                sw.WriteLine("            Ri_damping = " + myEquation_Tm.Ri_damping.ToString());
                sw.WriteLine("         EndSect  // EQUATION");
                myEquation_Tm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.SOLUTION_TECHNIQUE mySolution_Technique_Tm = myTurbulence_Module.solution_technique;
                sw.WriteLine("         [SOLUTION_TECHNIQUE]");
                sw.WriteLine("            Touched = " + mySolution_Technique_Tm.Touched.ToString());
                sw.WriteLine("            scheme_of_time_integration = " + mySolution_Technique_Tm.scheme_of_time_integration.ToString());
                sw.WriteLine("            scheme_of_space_discretization_horizontal = " + mySolution_Technique_Tm.scheme_of_space_discretization_horizontal.ToString());
                sw.WriteLine("            scheme_of_space_discretization_vertical = " + mySolution_Technique_Tm.scheme_of_space_discretization_vertical.ToString());
                sw.WriteLine("            method_of_space_discretization_horizontal = " + mySolution_Technique_Tm.method_of_space_discretization_horizontal.ToString());
                sw.WriteLine("         EndSect  // SOLUTION_TECHNIQUE");
                mySolution_Technique_Tm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.DIFFUSION myDiffusion_Tm = myTurbulence_Module.diffusion;
                sw.WriteLine("         [DIFFUSION]");
                sw.WriteLine("            Touched = " + myDiffusion_Tm.Touched.ToString());
                sw.WriteLine("            sigma_k_h = " + myDiffusion_Tm.sigma_k_h.ToString());
                sw.WriteLine("            sigma_e_h = " + myDiffusion_Tm.sigma_e_h.ToString());
                sw.WriteLine("            sigma_k = " + myDiffusion_Tm.sigma_k.ToString());
                sw.WriteLine("            sigma_e = " + myDiffusion_Tm.sigma_e.ToString());
                sw.WriteLine("         EndSect  // DIFFUSION");
                myDiffusion_Tm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.SOURCES mySources_Tm = myTurbulence_Module.sources;
                sw.WriteLine("         [SOURCES]");
                sw.WriteLine("            Touched = " + mySources_Tm.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + mySources_Tm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.SOURCES.SOURCE> mySource_Tm in mySources_Tm.source)
                {
                    sw.WriteLine("            [" + mySource_Tm.Key.ToString() + "]");
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.SOURCES.SOURCE.KINETIC_ENERGY myKinetic_Energy = mySource_Tm.Value.kinetic_energy;
                    sw.WriteLine("               [KINETIC_ENERGY]");
                    sw.WriteLine("                  Touched = " + myKinetic_Energy.Touched.ToString());
                    sw.WriteLine("                  type = " + myKinetic_Energy.type.ToString());
                    sw.WriteLine("                  format = " + myKinetic_Energy.format.ToString());
                    sw.WriteLine("                  constant_value = " + myKinetic_Energy.constant_value.ToString());
                    sw.WriteLine("                  file_name = " + myKinetic_Energy.file_name.ToString());
                    sw.WriteLine("                  item_number = " + myKinetic_Energy.item_number.ToString());
                    sw.WriteLine("                  item_name = " + myKinetic_Energy.item_name.ToString());
                    sw.WriteLine("                  type_of_soft_start = " + myKinetic_Energy.type_of_soft_start.ToString());
                    sw.WriteLine("                  soft_time_interval = " + myKinetic_Energy.soft_time_interval.ToString());
                    sw.WriteLine("                  reference_value = " + myKinetic_Energy.reference_value.ToString());
                    sw.WriteLine("                  type_of_time_interpolation = " + myKinetic_Energy.type_of_time_interpolation.ToString());
                    sw.WriteLine("               EndSect  // KINETIC_ENERGY");
                    myKinetic_Energy = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.SOURCES.SOURCE.DISSIPATION_OF_KINETIC_ENERGY myDissipation_Of_Kinetic_Energy = mySource_Tm.Value.dissipation_of_kinetic_energy;
                    sw.WriteLine("               [DISSIPATION_OF_KINETIC_ENERGY]");
                    sw.WriteLine("                  Touched = " + myDissipation_Of_Kinetic_Energy.Touched.ToString());
                    sw.WriteLine("                  type = " + myDissipation_Of_Kinetic_Energy.type.ToString());
                    sw.WriteLine("                  format = " + myDissipation_Of_Kinetic_Energy.format.ToString());
                    sw.WriteLine("                  constant_value = " + myDissipation_Of_Kinetic_Energy.constant_value.ToString());
                    sw.WriteLine("                  file_name = " + myDissipation_Of_Kinetic_Energy.file_name.ToString());
                    sw.WriteLine("                  item_number = " + myDissipation_Of_Kinetic_Energy.item_number.ToString());
                    sw.WriteLine("                  item_name = " + myDissipation_Of_Kinetic_Energy.item_name.ToString());
                    sw.WriteLine("                  type_of_soft_start = " + myDissipation_Of_Kinetic_Energy.type_of_soft_start.ToString());
                    sw.WriteLine("                  soft_time_interval = " + myDissipation_Of_Kinetic_Energy.soft_time_interval.ToString());
                    sw.WriteLine("                  reference_value = " + myDissipation_Of_Kinetic_Energy.reference_value.ToString());
                    sw.WriteLine("                  type_of_time_interpolation = " + myDissipation_Of_Kinetic_Energy.type_of_time_interpolation.ToString());
                    sw.WriteLine("               EndSect  // DISSIPATION_OF_KINETIC_ENERGY");
                    myDissipation_Of_Kinetic_Energy = null;
                    sw.WriteLine("");
                    sw.WriteLine("            EndSect  // " + mySource_Tm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("         EndSect  // SOURCES");
                mySources_Tm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.INITIAL_CONDITIONS myInitial_Conditions_Tm = myTurbulence_Module.initial_conditions;
                sw.WriteLine("         [INITIAL_CONDITIONS]");
                sw.WriteLine("            Touched = " + myInitial_Conditions_Tm.Touched.ToString());
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.INITIAL_CONDITIONS.KINETIC_ENERGY myInitial_Conditions_Kinetic_Energy_TM = myInitial_Conditions_Tm.kinetic_energy;
                sw.WriteLine("            [KINETIC_ENERGY]");
                sw.WriteLine("               Touched = " + myInitial_Conditions_Kinetic_Energy_TM.Touched.ToString());
                sw.WriteLine("               type = " + myInitial_Conditions_Kinetic_Energy_TM.type.ToString());
                sw.WriteLine("               format = " + myInitial_Conditions_Kinetic_Energy_TM.format.ToString());
                sw.WriteLine("               constant_value = " + myInitial_Conditions_Kinetic_Energy_TM.constant_value.ToString());
                sw.WriteLine("               file_name = " + myInitial_Conditions_Kinetic_Energy_TM.file_name.ToString());
                sw.WriteLine("               item_number = " + myInitial_Conditions_Kinetic_Energy_TM.item_number.ToString());
                sw.WriteLine("               item_name = " + myInitial_Conditions_Kinetic_Energy_TM.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myInitial_Conditions_Kinetic_Energy_TM.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myInitial_Conditions_Kinetic_Energy_TM.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myInitial_Conditions_Kinetic_Energy_TM.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myInitial_Conditions_Kinetic_Energy_TM.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // KINETIC_ENERGY");
                myInitial_Conditions_Kinetic_Energy_TM = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.INITIAL_CONDITIONS.DISSIPATION_OF_KINETIC_ENERGY myInitial_Conditions_Dissapation_of_Kinetic_energy = myInitial_Conditions_Tm.dissipation_of_kinetic_energy;
                sw.WriteLine("            [DISSIPATION_OF_KINETIC_ENERGY]");
                sw.WriteLine("               Touched = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.Touched.ToString());
                sw.WriteLine("               type = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.type.ToString());
                sw.WriteLine("               format = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.format.ToString());
                sw.WriteLine("               constant_value = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.constant_value.ToString());
                sw.WriteLine("               file_name = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.file_name.ToString());
                sw.WriteLine("               item_number = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.item_number.ToString());
                sw.WriteLine("               item_name = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.item_name.ToString());
                sw.WriteLine("               type_of_soft_start = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.type_of_soft_start.ToString());
                sw.WriteLine("               soft_time_interval = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.soft_time_interval.ToString());
                sw.WriteLine("               reference_value = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.reference_value.ToString());
                sw.WriteLine("               type_of_time_interpolation = " + myInitial_Conditions_Dissapation_of_Kinetic_energy.type_of_time_interpolation.ToString());
                sw.WriteLine("            EndSect  // DISSIPATION_OF_KINETIC_ENERGY");
                myInitial_Conditions_Dissapation_of_Kinetic_energy = null;
                sw.WriteLine("");
                sw.WriteLine("         EndSect  // INITIAL_CONDITIONS");
                myInitial_Conditions_Tm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.BOUNDARY_CONDITIONS myBoundary_Conditions_Tm = myTurbulence_Module.boundary_conditions;
                sw.WriteLine("         [BOUNDARY_CONDITIONS]");
                sw.WriteLine("            Touched = " + myBoundary_Conditions_Tm.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + myBoundary_Conditions_Tm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.BOUNDARY_CONDITIONS.CODE> myCode_tm in myBoundary_Conditions_Tm.code)
                {
                    sw.WriteLine("            [" + myCode_tm.Key.ToString() + "]");
                    if (myCode_tm.Key.ToString().Trim().ToUpper() == "CODE_1")
                    {
                        sw.WriteLine("               [KINETIC_ENERGY]");
                        sw.WriteLine("               EndSect  // KINETIC_ENERGY");
                        sw.WriteLine("");
                        sw.WriteLine("               [DISSIPATION_OF_KINETIC_ENERGY]");
                        sw.WriteLine("               EndSect  // DISSIPATION_OF_KINETIC_ENERGY");
                        sw.WriteLine("");
                    }
                    else
                    {
                        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.BOUNDARY_CONDITIONS.CODE.KINETIC_ENERGY myKinetic_Energy_tm_Bc = myCode_tm.Value.kinetic_energy;
                        sw.WriteLine("               [KINETIC_ENERGY]");
                        sw.WriteLine("                  identifier = " + myKinetic_Energy_tm_Bc.identifier.ToString());
                        sw.WriteLine("                  type = " + myKinetic_Energy_tm_Bc.type.ToString());
                        sw.WriteLine("                  type_interpolation_constrain = " + myKinetic_Energy_tm_Bc.type_interpolation_constrain.ToString());
                        sw.WriteLine("                  type_secondary = " + myKinetic_Energy_tm_Bc.type_secondary.ToString());
                        sw.WriteLine("                  type_of_vertical_profile = " + myKinetic_Energy_tm_Bc.type_of_vertical_profile.ToString());
                        sw.WriteLine("                  format = " + myKinetic_Energy_tm_Bc.format.ToString());
                        sw.WriteLine("                  constant_value = " + myKinetic_Energy_tm_Bc.constant_value.ToString());
                        sw.WriteLine("                  file_name = " + myKinetic_Energy_tm_Bc.file_name.ToString());
                        sw.WriteLine("                  item_number = " + myKinetic_Energy_tm_Bc.item_number.ToString());
                        sw.WriteLine("                  item_name = " + myKinetic_Energy_tm_Bc.item_name.ToString());
                        sw.WriteLine("                  type_of_soft_start = " + myKinetic_Energy_tm_Bc.type_of_soft_start.ToString());
                        sw.WriteLine("                  soft_time_interval = " + myKinetic_Energy_tm_Bc.soft_time_interval.ToString());
                        sw.WriteLine("                  reference_value = " + myKinetic_Energy_tm_Bc.reference_value.ToString());
                        sw.WriteLine("                  type_of_time_interpolation = " + myKinetic_Energy_tm_Bc.type_of_time_interpolation.ToString());
                        sw.WriteLine("                  type_of_space_interpolation = " + myKinetic_Energy_tm_Bc.type_of_space_interpolation.ToString());
                        sw.WriteLine("                  type_of_coriolis_correction = " + myKinetic_Energy_tm_Bc.type_of_coriolis_correction.ToString());
                        sw.WriteLine("                  type_of_wind_correction = " + myKinetic_Energy_tm_Bc.type_of_wind_correction.ToString());
                        sw.WriteLine("                  type_of_tilting = " + myKinetic_Energy_tm_Bc.type_of_tilting.ToString());
                        sw.WriteLine("                  type_of_tilting_point = " + myKinetic_Energy_tm_Bc.type_of_tilting_point.ToString());
                        sw.WriteLine("                  point_tilting = " + myKinetic_Energy_tm_Bc.point_tilting.ToString());
                        sw.WriteLine("                  type_of_radiation_stress_correction = " + myKinetic_Energy_tm_Bc.type_of_radiation_stress_correction.ToString());
                        sw.WriteLine("                  type_of_pressure_correction = " + myKinetic_Energy_tm_Bc.type_of_pressure_correction.ToString());
                        sw.WriteLine("                  type_of_radiation_stress_correction = " + myKinetic_Energy_tm_Bc.type_of_radiation_stress_correction2.ToString());
                        sw.WriteLine("               EndSect  // KINETIC_ENERGY");
                        myKinetic_Energy_tm_Bc = null;
                        sw.WriteLine("");
                        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.TURBULENCE_MODULE.BOUNDARY_CONDITIONS.CODE.DISSIPATION_OF_KINETIC_ENERGY myDissipation_Of_Kinetic_Energy_Tm_Bc = myCode_tm.Value.dissipation_of_kinetic_energy;
                        sw.WriteLine("               [DISSIPATION_OF_KINETIC_ENERGY]");
                        sw.WriteLine("                  identifier = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.identifier.ToString());
                        sw.WriteLine("                  type = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type.ToString());
                        sw.WriteLine("                  type_interpolation_constrain = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_interpolation_constrain.ToString());
                        sw.WriteLine("                  type_secondary = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_secondary.ToString());
                        sw.WriteLine("                  type_of_vertical_profile = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_vertical_profile.ToString());
                        sw.WriteLine("                  format = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.format.ToString());
                        sw.WriteLine("                  constant_value = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.constant_value.ToString());
                        sw.WriteLine("                  file_name = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.file_name.ToString());
                        sw.WriteLine("                  item_number = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.item_number.ToString());
                        sw.WriteLine("                  item_name = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.item_name.ToString());
                        sw.WriteLine("                  type_of_soft_start = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_soft_start.ToString());
                        sw.WriteLine("                  soft_time_interval = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.soft_time_interval.ToString());
                        sw.WriteLine("                  reference_value = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.reference_value.ToString());
                        sw.WriteLine("                  type_of_time_interpolation = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_time_interpolation.ToString());
                        sw.WriteLine("                  type_of_space_interpolation = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_space_interpolation.ToString());
                        sw.WriteLine("                  type_of_coriolis_correction = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_coriolis_correction.ToString());
                        sw.WriteLine("                  type_of_wind_correction = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_wind_correction.ToString());
                        sw.WriteLine("                  type_of_tilting = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_tilting.ToString());
                        sw.WriteLine("                  type_of_tilting_point = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_tilting_point.ToString());
                        sw.WriteLine("                  point_tilting = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.point_tilting.ToString());
                        sw.WriteLine("                  type_of_radiation_stress_correction = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_radiation_stress_correction.ToString());
                        sw.WriteLine("                  type_of_pressure_correction = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_pressure_correction.ToString());
                        sw.WriteLine("                  type_of_radiation_stress_correction = " + myDissipation_Of_Kinetic_Energy_Tm_Bc.type_of_radiation_stress_correction2.ToString());
                        sw.WriteLine("               EndSect  // DISSIPATION_OF_KINETIC_ENERGY");
                        myDissipation_Of_Kinetic_Energy_Tm_Bc = null;
                        sw.WriteLine("");
                    }
                    sw.WriteLine("            EndSect  // " + myCode_tm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("         EndSect  // BOUNDARY_CONDITIONS");
                myBoundary_Conditions_Tm = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // TURBULENCE_MODULE");
                myTurbulence_Module = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DECOUPLING myDecoupling = myHydrodynamic_Module.decoupling;
                sw.WriteLine("      [DECOUPLING]");
                sw.WriteLine("         Touched = " + myDecoupling.Touched.ToString());
                sw.WriteLine("         type = " + myDecoupling.type.ToString());
                sw.WriteLine("         file_name_flux = " + myDecoupling.file_name_flux.ToString());
                sw.WriteLine("         file_name_area = " + myDecoupling.file_name_area.ToString());
                sw.WriteLine("         file_name_volume = " + myDecoupling.file_name_volume.ToString());
                sw.WriteLine("         specification_file = " + myDecoupling.specification_file.ToString());
                sw.WriteLine("         first_time_step = " + myDecoupling.first_time_step.ToString());
                sw.WriteLine("         last_time_step = " + myDecoupling.last_time_step.ToString());
                sw.WriteLine("         time_step_frequency = " + myDecoupling.time_step_frequency.ToString());
                sw.WriteLine("      EndSect  // DECOUPLING");
                myDecoupling = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS myOutputs_Hd = myHydrodynamic_Module.outputs;
                sw.WriteLine("      [OUTPUTS]");
                sw.WriteLine("         Touched = " + myOutputs_Hd.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myOutputs_Hd.MzSEPfsListItemCount.ToString());
                sw.WriteLine("         number_of_outputs = " + myOutputs_Hd.number_of_outputs.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT> myOutput_Hd in myOutputs_Hd.output)
                {
                    sw.WriteLine("         [" + myOutput_Hd.Key.ToString() + "]");
                    sw.WriteLine("            Touched = " + myOutput_Hd.Value.Touched.ToString());
                    sw.WriteLine("            include = " + myOutput_Hd.Value.include.ToString());
                    sw.WriteLine("            title = " + myOutput_Hd.Value.title.ToString());
                    sw.WriteLine("            file_name = " + myOutput_Hd.Value.file_name.ToString());
                    sw.WriteLine("            type = " + myOutput_Hd.Value.type.ToString());
                    sw.WriteLine("            format = " + myOutput_Hd.Value.format.ToString());
                    sw.WriteLine("            flood_and_dry = " + myOutput_Hd.Value.flood_and_dry.ToString());
                    sw.WriteLine("            coordinate_type = " + myOutput_Hd.Value.coordinate_type.ToString());
                    sw.WriteLine("            zone = " + myOutput_Hd.Value.zone.ToString());
                    sw.WriteLine("            input_file_name = " + myOutput_Hd.Value.input_file_name.ToString());
                    sw.WriteLine("            input_format = " + myOutput_Hd.Value.input_format.ToString());
                    sw.WriteLine("            interpolation_type = " + myOutput_Hd.Value.interpolation_type.ToString());
                    sw.WriteLine("            first_time_step = " + myOutput_Hd.Value.first_time_step.ToString());
                    sw.WriteLine("            last_time_step = " + myOutput_Hd.Value.last_time_step.ToString());
                    sw.WriteLine("            time_step_frequency = " + myOutput_Hd.Value.time_step_frequency.ToString());
                    sw.WriteLine("            number_of_points = " + myOutput_Hd.Value.number_of_points.ToString());
                    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT.POINT> myPoint_Out_Hd in myOutput_Hd.Value.point)
                    {
                        sw.WriteLine("            [" + myPoint_Out_Hd.Key.ToString() + "]");
                        sw.WriteLine("               name = " + myPoint_Out_Hd.Value.name.ToString());
                        sw.WriteLine("               x = " + myPoint_Out_Hd.Value.x.ToString());
                        sw.WriteLine("               y = " + myPoint_Out_Hd.Value.y.ToString());
                        sw.WriteLine("               z = " + myPoint_Out_Hd.Value.z.ToString());
                        sw.WriteLine("               layer = " + myPoint_Out_Hd.Value.layer.ToString());
                        sw.WriteLine("            EndSect  // " + myPoint_Out_Hd.Key.ToString());
                        sw.WriteLine("");
                    }
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT.LINE myLine_Out_Hd = myOutput_Hd.Value.line;
                    sw.WriteLine("            [LINE]");
                    sw.WriteLine("               npoints = " + myLine_Out_Hd.npoints.ToString());
                    sw.WriteLine("               x_first = " + myLine_Out_Hd.x_first.ToString());
                    sw.WriteLine("               y_first = " + myLine_Out_Hd.y_first.ToString());
                    sw.WriteLine("               z_first = " + myLine_Out_Hd.z_first.ToString());
                    sw.WriteLine("               x_last = " + myLine_Out_Hd.x_last.ToString());
                    sw.WriteLine("               y_last = " + myLine_Out_Hd.y_last.ToString());
                    sw.WriteLine("               z_last = " + myLine_Out_Hd.z_last.ToString());
                    sw.WriteLine("            EndSect  // LINE");
                    myLine_Out_Hd = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT.AREA myArea_Out_Hd = myOutput_Hd.Value.area;
                    sw.WriteLine("            [AREA]");
                    sw.WriteLine("               number_of_points = " + myArea_Out_Hd.number_of_points.ToString());
                    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT.AREA.POINT> myPoint_Area_Out_Hd in myArea_Out_Hd.point)
                    {
                        sw.WriteLine("               [" + myPoint_Area_Out_Hd.Key.ToString() + "]");
                        sw.WriteLine("                  x = " + myPoint_Area_Out_Hd.Value.x.ToString());
                        sw.WriteLine("                  y = " + myPoint_Area_Out_Hd.Value.y.ToString());
                        sw.WriteLine("               EndSect  // " + myPoint_Area_Out_Hd.Key.ToString());
                        sw.WriteLine("");
                    }
                    sw.WriteLine("               layer_min = " + myArea_Out_Hd.layer_min.ToString());
                    sw.WriteLine("               layer_max = " + myArea_Out_Hd.layer_max.ToString());
                    sw.WriteLine("            EndSect  // AREA");
                    myArea_Out_Hd = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT.PARAMETERS_2D myParameters_2d_Hd = myOutput_Hd.Value.parameters_2d;
                    sw.WriteLine("            [PARAMETERS_2D]");
                    sw.WriteLine("               Touched = " + myParameters_2d_Hd.Touched.ToString());
                    sw.WriteLine("               SURFACE_ELEVATION = " + myParameters_2d_Hd.SURFACE_ELEVATION.ToString());
                    sw.WriteLine("               STILL_WATER_DEPTH = " + myParameters_2d_Hd.STILL_WATER_DEPTH.ToString());
                    sw.WriteLine("               TOTAL_WATER_DEPTH = " + myParameters_2d_Hd.TOTAL_WATER_DEPTH.ToString());
                    sw.WriteLine("               U_VELOCITY = " + myParameters_2d_Hd.U_VELOCITY.ToString());
                    sw.WriteLine("               V_VELOCITY = " + myParameters_2d_Hd.V_VELOCITY.ToString());
                    sw.WriteLine("               P_FLUX = " + myParameters_2d_Hd.P_FLUX.ToString());
                    sw.WriteLine("               Q_FLUX = " + myParameters_2d_Hd.Q_FLUX.ToString());
                    sw.WriteLine("               DENSITY = " + myParameters_2d_Hd.DENSITY.ToString());
                    sw.WriteLine("               TEMPERATURE = " + myParameters_2d_Hd.TEMPERATURE.ToString());
                    sw.WriteLine("               SALINITY = " + myParameters_2d_Hd.SALINITY.ToString());
                    sw.WriteLine("               CURRENT_SPEED = " + myParameters_2d_Hd.CURRENT_SPEED.ToString());
                    sw.WriteLine("               CURRENT_DIRECTION = " + myParameters_2d_Hd.CURRENT_DIRECTION.ToString());
                    sw.WriteLine("               WIND_U_VELOCITY = " + myParameters_2d_Hd.WIND_U_VELOCITY.ToString());
                    sw.WriteLine("               WIND_V_VELOCITY = " + myParameters_2d_Hd.WIND_V_VELOCITY.ToString());
                    sw.WriteLine("               AIR_PRESSURE = " + myParameters_2d_Hd.AIR_PRESSURE.ToString());
                    sw.WriteLine("               PRECIPITATION = " + myParameters_2d_Hd.PRECIPITATION.ToString());
                    sw.WriteLine("               EVAPORATION = " + myParameters_2d_Hd.EVAPORATION.ToString());
                    sw.WriteLine("               DRAG_COEFFICIENT = " + myParameters_2d_Hd.DRAG_COEFFICIENT.ToString());
                    sw.WriteLine("               EDDY_VISCOSITY = " + myParameters_2d_Hd.EDDY_VISCOSITY.ToString());
                    sw.WriteLine("               CFL_NUMBER = " + myParameters_2d_Hd.CFL_NUMBER.ToString());
                    sw.WriteLine("               CONVERGENCE_ANGLE = " + myParameters_2d_Hd.CONVERGENCE_ANGLE.ToString());
                    sw.WriteLine("               AREA = " + myParameters_2d_Hd.AREA.ToString());
                    sw.WriteLine("            EndSect  // PARAMETERS_2D");
                    myParameters_2d_Hd = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT.PARAMETERS_3D myParameters_3d_Hd = myOutput_Hd.Value.parameters_3d;
                    sw.WriteLine("            [PARAMETERS_3D]");
                    sw.WriteLine("               Touched = " + myParameters_3d_Hd.Touched.ToString());
                    sw.WriteLine("               U_VELOCITY = " + myParameters_3d_Hd.U_VELOCITY.ToString());
                    sw.WriteLine("               V_VELOCITY = " + myParameters_3d_Hd.V_VELOCITY.ToString());
                    sw.WriteLine("               W_VELOCITY = " + myParameters_3d_Hd.W_VELOCITY.ToString());
                    sw.WriteLine("               WS_VELOCITY = " + myParameters_3d_Hd.WS_VELOCITY.ToString());
                    sw.WriteLine("               DENSITY = " + myParameters_3d_Hd.DENSITY.ToString());
                    sw.WriteLine("               TEMPERATURE = " + myParameters_3d_Hd.TEMPERATURE.ToString());
                    sw.WriteLine("               SALINITY = " + myParameters_3d_Hd.SALINITY.ToString());
                    sw.WriteLine("               TURBULENT_KINETIC_ENERGY = " + myParameters_3d_Hd.TURBULENT_KINETIC_ENERGY.ToString());
                    sw.WriteLine("               DISSIPATION_OF_TKE = " + myParameters_3d_Hd.DISSIPATION_OF_TKE.ToString());
                    sw.WriteLine("               CURRENT_SPEED = " + myParameters_3d_Hd.CURRENT_SPEED.ToString());
                    sw.WriteLine("               CURRENT_DIRECTION_HORIZONTAL = " + myParameters_3d_Hd.CURRENT_DIRECTION_HORIZONTAL.ToString());
                    sw.WriteLine("               CURRENT_DIRECTION_VERTICAL = " + myParameters_3d_Hd.CURRENT_DIRECTION_VERTICAL.ToString());
                    sw.WriteLine("               HORIZONTAL_EDDY_VISCOSITY = " + myParameters_3d_Hd.HORIZONTAL_EDDY_VISCOSITY.ToString());
                    sw.WriteLine("               VERTICAL_EDDY_VISCOSITY = " + myParameters_3d_Hd.VERTICAL_EDDY_VISCOSITY.ToString());
                    sw.WriteLine("               CFL_NUMBER = " + myParameters_3d_Hd.CFL_NUMBER.ToString());
                    sw.WriteLine("               VOLUME = " + myParameters_3d_Hd.VOLUME.ToString());
                    sw.WriteLine("            EndSect  // PARAMETERS_3D");
                    myParameters_3d_Hd = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT.DISCHARGE myDischarge_Hd = myOutput_Hd.Value.discharge;
                    sw.WriteLine("            [DISCHARGE]");
                    sw.WriteLine("               Touched = " + myDischarge_Hd.Touched.ToString());
                    sw.WriteLine("               DISCHARGE = " + myDischarge_Hd.discharge.ToString());
                    sw.WriteLine("               ACCUMULATED_DISCHARGE = " + myDischarge_Hd.ACCUMULATED_DISCHARGE.ToString());
                    sw.WriteLine("               FLOW = " + myDischarge_Hd.FLOW.ToString());
                    sw.WriteLine("               TEMPERATURE = " + myDischarge_Hd.TEMPERATURE.ToString());
                    sw.WriteLine("               SALINITY = " + myDischarge_Hd.SALINITY.ToString());
                    sw.WriteLine("            EndSect  // DISCHARGE");
                    myDischarge_Hd = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.OUTPUTS.OUTPUT.MASSBUDGET myMassbuget_Hd = myOutput_Hd.Value.massbudget;
                    sw.WriteLine("            [MASSBUDGET]");
                    sw.WriteLine("               Touched = " + myMassbuget_Hd.Touched.ToString());
                    sw.WriteLine("               DISCHARGE = " + myMassbuget_Hd.DISCHARGE.ToString());
                    sw.WriteLine("               ACCUMULATED_DISCHARGE = " + myMassbuget_Hd.ACCUMULATED_DISCHARGE.ToString());
                    sw.WriteLine("               FLOW = " + myMassbuget_Hd.FLOW.ToString());
                    sw.WriteLine("               TEMPERATURE = " + myMassbuget_Hd.TEMPERATURE.ToString());
                    sw.WriteLine("               SALINITY = " + myMassbuget_Hd.SALINITY.ToString());
                    sw.WriteLine("            EndSect  // MASSBUDGET");
                    myMassbuget_Hd = null; ;
                    sw.WriteLine("");
                    sw.WriteLine("         EndSect  // " + myOutput_Hd.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // OUTPUTS");
                sw.WriteLine("");
                sw.WriteLine("   EndSect  // HYDRODYNAMIC_MODULE");
                myOutputs_Hd = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE myTransport_Module = myFemEnginHD.transport_module;
                sw.WriteLine("   [TRANSPORT_MODULE]");
                sw.WriteLine("      mode = " + myTransport_Module.mode.ToString());
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.EQUATION myTransport_Module_Equation = myTransport_Module.equation;
                sw.WriteLine("      [EQUATION]");
                sw.WriteLine("         formulation = " + myTransport_Module_Equation.formulation.ToString());
                sw.WriteLine("      EndSect  // EQUATION");
                myTransport_Module_Equation = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.TIME myTransport_Module_Time = myTransport_Module.time;
                sw.WriteLine("      [TIME]");
                sw.WriteLine("         start_time_step = " + myTransport_Module_Time.start_time_step.ToString());
                sw.WriteLine("         time_step_factor = " + myTransport_Module_Time.time_step_factor.ToString());
                sw.WriteLine("      EndSect  // TIME");
                myTransport_Module_Time = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.SPACE mySpace_Trm = myTransport_Module.space_TRM;
                sw.WriteLine("      [SPACE]");
                sw.WriteLine("         number_of_2D_mesh_geometry = " + mySpace_Trm.number_of_2D_mesh_geometry.ToString());
                sw.WriteLine("         number_of_2D_mesh_velocity = " + mySpace_Trm.number_of_2D_mesh_velocity.ToString());
                sw.WriteLine("         number_of_2D_mesh_concentration = " + mySpace_Trm.number_of_2D_mesh_concentration.ToString());
                sw.WriteLine("         number_of_3D_mesh_geometry = " + mySpace_Trm.number_of_3D_mesh_geometry.ToString());
                sw.WriteLine("         number_of_3D_mesh_velocity = " + mySpace_Trm.number_of_3D_mesh_velocity.ToString());
                sw.WriteLine("         number_of_3D_mesh_concentration = " + mySpace_Trm.number_of_3D_mesh_concentration.ToString());
                sw.WriteLine("      EndSect  // SPACE");
                mySpace_Trm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.COMPONENTS myComponents_Trn = myTransport_Module.components;
                sw.WriteLine("      [COMPONENTS]");
                sw.WriteLine("         Touched = " + myComponents_Trn.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myComponents_Trn.MzSEPfsListItemCount.ToString());
                sw.WriteLine("         number_of_components = " + myComponents_Trn.number_of_components.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.COMPONENTS.COMPONENT> myComponent_Trn in myComponents_Trn.component)
                {
                    sw.WriteLine("         [" + myComponent_Trn.Key.ToString() + "]");
                    sw.WriteLine("            Touched = " + myComponent_Trn.Value.Touched.ToString());
                    sw.WriteLine("            name = " + myComponent_Trn.Value.name.ToString());
                    sw.WriteLine("            type = " + myComponent_Trn.Value.type.ToString());
                    sw.WriteLine("            dimension = " + myComponent_Trn.Value.dimension.ToString());
                    sw.WriteLine("            description = " + myComponent_Trn.Value.description.ToString());
                    sw.WriteLine("            EUM_type = " + myComponent_Trn.Value.EUM_type.ToString());
                    sw.WriteLine("            EUM_unit = " + myComponent_Trn.Value.EUM_unit.ToString());
                    sw.WriteLine("            unit = " + myComponent_Trn.Value.unit.ToString());
                    sw.WriteLine("            minimum_value = " + myComponent_Trn.Value.minimum_value.ToString());
                    sw.WriteLine("            maximum_value = " + myComponent_Trn.Value.maximum_value.ToString());
                    sw.WriteLine("         EndSect  // " + myComponent_Trn.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // COMPONENTS");
                myComponents_Trn = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.SOLUTION_TECHNIQUE mySolution_Technique_Trm = myTransport_Module.solution_technique;
                sw.WriteLine("      [SOLUTION_TECHNIQUE]");
                sw.WriteLine("         Touched = " + mySolution_Technique_Trm.Touched.ToString());
                sw.WriteLine("         scheme_of_time_integration = " + mySolution_Technique_Trm.scheme_of_time_integration.ToString());
                sw.WriteLine("         scheme_of_space_discretization_horizontal = " + mySolution_Technique_Trm.scheme_of_space_discretization_horizontal.ToString());
                sw.WriteLine("         scheme_of_space_discretization_vertical = " + mySolution_Technique_Trm.scheme_of_space_discretization_vertical.ToString());
                sw.WriteLine("         method_of_space_discretization_horizontal = " + mySolution_Technique_Trm.method_of_space_discretization_horizontal.ToString());
                sw.WriteLine("      EndSect  // SOLUTION_TECHNIQUE");
                mySolution_Technique_Trm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.HYDRODYNAMIC_CONDITIONS myHydrodynamic_Contitions = myTransport_Module.hydrodynamic_conditions;
                sw.WriteLine("      [HYDRODYNAMIC_CONDITIONS]");
                sw.WriteLine("         Touched = " + myHydrodynamic_Contitions.Touched.ToString());
                sw.WriteLine("         type = " + myHydrodynamic_Contitions.type.ToString());
                sw.WriteLine("         format = " + myHydrodynamic_Contitions.format.ToString());
                sw.WriteLine("         surface_elevation_constant = " + myHydrodynamic_Contitions.surface_elevation_constant.ToString());
                sw.WriteLine("         u_velocity_constant = " + myHydrodynamic_Contitions.u_velocity_constant.ToString());
                sw.WriteLine("         v_velocity_constant = " + myHydrodynamic_Contitions.v_velocity_constant.ToString());
                sw.WriteLine("         w_velocity_constant = " + myHydrodynamic_Contitions.w_velocity_constant.ToString());
                sw.WriteLine("         file_name = " + myHydrodynamic_Contitions.file_name.ToString());
                sw.WriteLine("      EndSect  // HYDRODYNAMIC_CONDITIONS");
                myHydrodynamic_Contitions = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION myDispersion_Trm = myTransport_Module.dispersion;
                sw.WriteLine("      [DISPERSION]");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION.HORIZONTAL_DISPERSION myHorizontal_Dispersion_Trm = myDispersion_Trm.horizontal_dispersion;
                sw.WriteLine("         [HORIZONTAL_DISPERSION]");
                sw.WriteLine("            Touched = " + myHorizontal_Dispersion_Trm.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + myHorizontal_Dispersion_Trm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION.HORIZONTAL_DISPERSION.COMPONENT> myHorizontal_Dispersion_Component_Trm in myHorizontal_Dispersion_Trm.component)
                {
                    sw.WriteLine("            [" + myHorizontal_Dispersion_Component_Trm.Key.ToString() + "]");
                    sw.WriteLine("               Touched = " + myHorizontal_Dispersion_Component_Trm.Value.Touched.ToString());
                    sw.WriteLine("               type = " + myHorizontal_Dispersion_Component_Trm.Value.type.ToString());
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION.HORIZONTAL_DISPERSION.COMPONENT.SCALED_EDDY_VISCOSITY myScaled_Eddy_Viscosity_Trm = myHorizontal_Dispersion_Component_Trm.Value.scaled_eddy_viscosity;
                    sw.WriteLine("               [SCALED_EDDY_VISCOSITY]");
                    sw.WriteLine("                  format = " + myScaled_Eddy_Viscosity_Trm.format.ToString());
                    sw.WriteLine("                  sigma = " + myScaled_Eddy_Viscosity_Trm.sigma.ToString());
                    sw.WriteLine("                  file_name = " + myScaled_Eddy_Viscosity_Trm.file_name.ToString());
                    sw.WriteLine("                  item_number = " + myScaled_Eddy_Viscosity_Trm.item_number.ToString());
                    sw.WriteLine("                  item_name = " + myScaled_Eddy_Viscosity_Trm.item_name.ToString());
                    sw.WriteLine("               EndSect  // SCALED_EDDY_VISCOSITY");
                    myScaled_Eddy_Viscosity_Trm = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION.HORIZONTAL_DISPERSION.COMPONENT.DISPERSION_COEFFICIENT myDispersion_Coefficient_Trm = myHorizontal_Dispersion_Component_Trm.Value.dispersion_coefficient;
                    sw.WriteLine("               [DISPERSION_COEFFICIENT]");
                    sw.WriteLine("                  format = " + myDispersion_Coefficient_Trm.format.ToString());
                    sw.WriteLine("                  constant_value = " + myDispersion_Coefficient_Trm.constant_value.ToString());
                    sw.WriteLine("                  file_name = " + myDispersion_Coefficient_Trm.file_name.ToString());
                    sw.WriteLine("                  item_number = " + myDispersion_Coefficient_Trm.item_number.ToString());
                    sw.WriteLine("                  item_name = " + myDispersion_Coefficient_Trm.item_name.ToString());
                    sw.WriteLine("               EndSect  // DISPERSION_COEFFICIENT");
                    myDispersion_Coefficient_Trm = null;
                    sw.WriteLine("");
                    sw.WriteLine("            EndSect  // " + myHorizontal_Dispersion_Component_Trm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("         EndSect  // HORIZONTAL_DISPERSION");
                myHorizontal_Dispersion_Trm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION.VERTICAL_DISPERSION myVertical_Dispersion_Trm = myDispersion_Trm.vertical_dispersion;
                sw.WriteLine("         [VERTICAL_DISPERSION]");
                sw.WriteLine("            Touched = " + myVertical_Dispersion_Trm.Touched.ToString());
                sw.WriteLine("            MzSEPfsListItemCount = " + myVertical_Dispersion_Trm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION.VERTICAL_DISPERSION.COMPONENT> myVertical_Dispersion_Component_Trm in myVertical_Dispersion_Trm.component)
                {
                    sw.WriteLine("            [" + myVertical_Dispersion_Component_Trm.Key.ToString() + "]");
                    sw.WriteLine("               Touched = " + myVertical_Dispersion_Component_Trm.Value.Touched.ToString());
                    sw.WriteLine("               type = " + myVertical_Dispersion_Component_Trm.Value.type.ToString());
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION.VERTICAL_DISPERSION.COMPONENT.SCALED_EDDY_VISCOSITY myScaled_Eddy_Viscosity_Trm = myVertical_Dispersion_Component_Trm.Value.scaled_eddy_viscosity;
                    sw.WriteLine("               [SCALED_EDDY_VISCOSITY]");
                    sw.WriteLine("                  format = " + myScaled_Eddy_Viscosity_Trm.format.ToString());
                    sw.WriteLine("                  sigma = " + myScaled_Eddy_Viscosity_Trm.sigma.ToString());
                    sw.WriteLine("                  file_name = " + myScaled_Eddy_Viscosity_Trm.file_name.ToString());
                    sw.WriteLine("                  item_number = " + myScaled_Eddy_Viscosity_Trm.item_number.ToString());
                    sw.WriteLine("                  item_name = " + myScaled_Eddy_Viscosity_Trm.item_name.ToString());
                    sw.WriteLine("               EndSect  // SCALED_EDDY_VISCOSITY");
                    myScaled_Eddy_Viscosity_Trm = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DISPERSION.VERTICAL_DISPERSION.COMPONENT.DISPERSION_COEFFICIENT myDispersion_Coefficient_Trm = myVertical_Dispersion_Component_Trm.Value.dispersion_coefficient;
                    sw.WriteLine("               [DISPERSION_COEFFICIENT]");
                    sw.WriteLine("                  format = " + myDispersion_Coefficient_Trm.format.ToString());
                    sw.WriteLine("                  constant_value = " + myDispersion_Coefficient_Trm.constant_value.ToString());
                    sw.WriteLine("                  file_name = " + myDispersion_Coefficient_Trm.file_name.ToString());
                    sw.WriteLine("                  item_number = " + myDispersion_Coefficient_Trm.item_number.ToString());
                    sw.WriteLine("                  item_name = " + myDispersion_Coefficient_Trm.item_name.ToString());
                    sw.WriteLine("               EndSect  // DISPERSION_COEFFICIENT");
                    myDispersion_Coefficient_Trm = null;
                    sw.WriteLine("");
                    sw.WriteLine("            EndSect  // " + myVertical_Dispersion_Component_Trm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("         EndSect  // VERTICAL_DISPERSION");
                myVertical_Dispersion_Trm = null;
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // DISPERSION");
                myDispersion_Trm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DECAY myDecay = myTransport_Module.decay;
                sw.WriteLine("      [DECAY]");
                sw.WriteLine("         Touched = " + myDecay.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myDecay.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.DECAY.COMPONENT> myComponent_Decay in myDecay.component)
                {
                    sw.WriteLine("         [" + myComponent_Decay.Key.ToString() + "]");
                    sw.WriteLine("            Touched = " + myComponent_Decay.Value.Touched.ToString());
                    sw.WriteLine("            type = " + myComponent_Decay.Value.type.ToString());
                    sw.WriteLine("            format = " + myComponent_Decay.Value.format.ToString());
                    sw.WriteLine("            constant_value = " + myComponent_Decay.Value.constant_value.ToString());
                    sw.WriteLine("            file_name = " + myComponent_Decay.Value.file_name.ToString());
                    sw.WriteLine("            item_number = " + myComponent_Decay.Value.item_number.ToString());
                    sw.WriteLine("            item_name = " + myComponent_Decay.Value.item_name.ToString());
                    sw.WriteLine("            type_of_soft_start = " + myComponent_Decay.Value.type_of_soft_start.ToString());
                    sw.WriteLine("            soft_time_interval = " + myComponent_Decay.Value.soft_time_interval.ToString());
                    sw.WriteLine("            reference_value = " + myComponent_Decay.Value.reference_value.ToString());
                    sw.WriteLine("            type_of_time_interpolation = " + myComponent_Decay.Value.type_of_time_interpolation.ToString());
                    sw.WriteLine("         EndSect  // " + myComponent_Decay.Key.ToString());
                }
                sw.WriteLine("");
                sw.WriteLine("      EndSect  // DECAY");
                myDecay = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.PRECIPITATION_EVAPORATION myPrecipitation_Evaporation_Trm = myTransport_Module.precipitation_evaporation;
                sw.WriteLine("      [PRECIPITATION_EVAPORATION]");
                sw.WriteLine("         Touched = " + myPrecipitation_Evaporation_Trm.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myPrecipitation_Evaporation_Trm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.PRECIPITATION_EVAPORATION.COMPONENT> myComponent_Pe_Trm in myPrecipitation_Evaporation_Trm.component)
                {
                    sw.WriteLine("         [" + myComponent_Pe_Trm.Key.ToString() + "]");
                    sw.WriteLine("            Touched = " + myComponent_Pe_Trm.Value.Touched.ToString());
                    sw.WriteLine("            type_of_precipitation = " + myComponent_Pe_Trm.Value.type_of_precipitation.ToString());
                    sw.WriteLine("            type_of_evaporation = " + myComponent_Pe_Trm.Value.type_of_evaporation.ToString());
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.PRECIPITATION_EVAPORATION.COMPONENT.PRECIPITATION myComponent_Pe_Trm_Precipitation = myComponent_Pe_Trm.Value.precipitation;
                    sw.WriteLine("            [PRECIPITATION]");
                    sw.WriteLine("               Touched = " + myComponent_Pe_Trm_Precipitation.Touched.ToString());
                    sw.WriteLine("               type = " + myComponent_Pe_Trm_Precipitation.type.ToString());
                    sw.WriteLine("               format = " + myComponent_Pe_Trm_Precipitation.format.ToString());
                    sw.WriteLine("               constant_value = " + myComponent_Pe_Trm_Precipitation.constant_value.ToString());
                    sw.WriteLine("               file_name = " + myComponent_Pe_Trm_Precipitation.file_name.ToString());
                    sw.WriteLine("               item_number = " + myComponent_Pe_Trm_Precipitation.item_number.ToString());
                    sw.WriteLine("               item_name = " + myComponent_Pe_Trm_Precipitation.item_name.ToString());
                    sw.WriteLine("               type_of_soft_start = " + myComponent_Pe_Trm_Precipitation.type_of_soft_start.ToString());
                    sw.WriteLine("               soft_time_interval = " + myComponent_Pe_Trm_Precipitation.soft_time_interval.ToString());
                    sw.WriteLine("               reference_value = " + myComponent_Pe_Trm_Precipitation.reference_value.ToString());
                    sw.WriteLine("               type_of_time_interpolation = " + myComponent_Pe_Trm_Precipitation.type_of_time_interpolation.ToString());
                    sw.WriteLine("            EndSect  // PRECIPITATION");
                    myComponent_Pe_Trm_Precipitation = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.PRECIPITATION_EVAPORATION.COMPONENT.EVAPORATION myComponent_Pe_Trm_Evaporation = myComponent_Pe_Trm.Value.evaporation;
                    sw.WriteLine("            [EVAPORATION]");
                    sw.WriteLine("               Touched = " + myComponent_Pe_Trm_Evaporation.Touched.ToString());
                    sw.WriteLine("               type = " + myComponent_Pe_Trm_Evaporation.type.ToString());
                    sw.WriteLine("               format = " + myComponent_Pe_Trm_Evaporation.format.ToString());
                    sw.WriteLine("               constant_value = " + myComponent_Pe_Trm_Evaporation.constant_value.ToString());
                    sw.WriteLine("               file_name = " + myComponent_Pe_Trm_Evaporation.file_name.ToString());
                    sw.WriteLine("               item_number = " + myComponent_Pe_Trm_Evaporation.item_number.ToString());
                    sw.WriteLine("               item_name = " + myComponent_Pe_Trm_Evaporation.item_name.ToString());
                    sw.WriteLine("               type_of_soft_start = " + myComponent_Pe_Trm_Evaporation.type_of_soft_start.ToString());
                    sw.WriteLine("               soft_time_interval = " + myComponent_Pe_Trm_Evaporation.soft_time_interval.ToString());
                    sw.WriteLine("               reference_value = " + myComponent_Pe_Trm_Evaporation.reference_value.ToString());
                    sw.WriteLine("               type_of_time_interpolation = " + myComponent_Pe_Trm_Evaporation.type_of_time_interpolation.ToString());
                    sw.WriteLine("            EndSect  // EVAPORATION");
                    myComponent_Pe_Trm_Evaporation = null;
                    sw.WriteLine("");
                    sw.WriteLine("         EndSect  // " + myComponent_Pe_Trm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // PRECIPITATION_EVAPORATION");
                myPrecipitation_Evaporation_Trm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.SOURCES mySources_Trm = myTransport_Module.sources;
                sw.WriteLine("      [SOURCES]");
                sw.WriteLine("         Touched = " + mySources_Trm.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + mySources_Trm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.SOURCES.SOURCE> mySource_Trm in mySources_Trm.source)
                {
                    sw.WriteLine("         [" + mySource_Trm.Key.ToString() + "]");
                    sw.WriteLine("            Touched = " + mySource_Trm.Value.Touched.ToString());
                    sw.WriteLine("            MzSEPfsListItemCount = " + mySource_Trm.Value.MzSEPfsListItemCount.ToString());
                    sw.WriteLine("            name = " + mySource_Trm.Value.name.ToString());
                    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.SOURCES.SOURCE.COMPONENT> myComponent_Source_Trm in mySource_Trm.Value.component)
                    {
                        sw.WriteLine("            [" + myComponent_Source_Trm.Key.ToString() + "]");
                        sw.WriteLine("               type_of_component = " + myComponent_Source_Trm.Value.type_of_component.ToString());
                        sw.WriteLine("               type = " + myComponent_Source_Trm.Value.type.ToString());
                        sw.WriteLine("               format = " + myComponent_Source_Trm.Value.format.ToString());
                        sw.WriteLine("               constant_value = " + myComponent_Source_Trm.Value.constant_value.ToString());
                        sw.WriteLine("               file_name = " + myComponent_Source_Trm.Value.file_name.ToString());
                        sw.WriteLine("               item_number = " + myComponent_Source_Trm.Value.item_number.ToString());
                        sw.WriteLine("               item_name = " + myComponent_Source_Trm.Value.item_name.ToString());
                        sw.WriteLine("               type_of_soft_start = " + myComponent_Source_Trm.Value.type_of_soft_start.ToString());
                        sw.WriteLine("               soft_time_interval = " + myComponent_Source_Trm.Value.soft_time_interval.ToString());
                        sw.WriteLine("               reference_value = " + myComponent_Source_Trm.Value.reference_value.ToString());
                        sw.WriteLine("               type_of_time_interpolation = " + myComponent_Source_Trm.Value.type_of_time_interpolation.ToString());
                        sw.WriteLine("            EndSect  // " + myComponent_Source_Trm.Key.ToString());
                        sw.WriteLine("");
                    }
                    sw.WriteLine("         EndSect  // " + mySource_Trm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // SOURCES");
                mySources_Trm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.INITIAL_CONDITIONS myInitial_Conditions_Trm = myTransport_Module.initial_conditions;
                sw.WriteLine("      [INITIAL_CONDITIONS]");
                sw.WriteLine("         Touched = " + myInitial_Conditions_Trm.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myInitial_Conditions_Trm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.INITIAL_CONDITIONS.COMPONENT> myComponent_Init_Cond_Trm in myInitial_Conditions_Trm.component)
                {
                    sw.WriteLine("         [" + myComponent_Init_Cond_Trm.Key.ToString() + "]");
                    sw.WriteLine("            Touched = " + myComponent_Init_Cond_Trm.Value.Touched.ToString());
                    sw.WriteLine("            type = " + myComponent_Init_Cond_Trm.Value.type.ToString());
                    sw.WriteLine("            format = " + myComponent_Init_Cond_Trm.Value.format.ToString());
                    sw.WriteLine("            constant_value = " + myComponent_Init_Cond_Trm.Value.constant_value.ToString());
                    sw.WriteLine("            file_name = " + myComponent_Init_Cond_Trm.Value.file_name.ToString());
                    sw.WriteLine("            item_number = " + myComponent_Init_Cond_Trm.Value.item_number.ToString());
                    sw.WriteLine("            item_name = " + myComponent_Init_Cond_Trm.Value.item_name.ToString());
                    sw.WriteLine("            type_of_soft_start = " + myComponent_Init_Cond_Trm.Value.type_of_soft_start.ToString());
                    sw.WriteLine("            soft_time_interval = " + myComponent_Init_Cond_Trm.Value.soft_time_interval.ToString());
                    sw.WriteLine("            reference_value = " + myComponent_Init_Cond_Trm.Value.reference_value.ToString());
                    sw.WriteLine("            type_of_time_interpolation = " + myComponent_Init_Cond_Trm.Value.type_of_time_interpolation.ToString());
                    sw.WriteLine("         EndSect  // " + myComponent_Init_Cond_Trm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // INITIAL_CONDITIONS");
                myInitial_Conditions_Trm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.BOUNDARY_CONDITIONS myBoundary_Conditions_Trm = myTransport_Module.boundary_conditions;
                sw.WriteLine("      [BOUNDARY_CONDITIONS]");
                sw.WriteLine("         Touched = " + myBoundary_Conditions_Trm.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myBoundary_Conditions_Trm.MzSEPfsListItemCount.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.BOUNDARY_CONDITIONS.CODE> myCode_Bc_Trm in myBoundary_Conditions_Trm.code)
                {
                    sw.WriteLine("         [" + myCode_Bc_Trm.Key.ToString() + "]");
                    if (myCode_Bc_Trm.Key.ToString().Trim().ToUpper() == "CODE_1")
                    {
                    }
                    else
                    {
                        sw.WriteLine("            Touched = " + myCode_Bc_Trm.Value.Touched.ToString());
                        sw.WriteLine("            MzSEPfsListItemCount = " + myCode_Bc_Trm.Value.MzSEPfsListItemCount.ToString());
                        foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.BOUNDARY_CONDITIONS.CODE.COMPONENT> myComonent_Bc_Trm in myCode_Bc_Trm.Value.component)
                        {
                            sw.WriteLine("            [" + myComonent_Bc_Trm.Key.ToString() + "]");
                            sw.WriteLine("               Touched = " + myComonent_Bc_Trm.Value.Touched.ToString());
                            sw.WriteLine("               type = " + myComonent_Bc_Trm.Value.type.ToString());
                            sw.WriteLine("               format = " + myComonent_Bc_Trm.Value.format.ToString());
                            sw.WriteLine("               constant_value = " + myComonent_Bc_Trm.Value.constant_value.ToString());
                            sw.WriteLine("               file_name = " + myComonent_Bc_Trm.Value.file_name.ToString());
                            sw.WriteLine("               item_number = " + myComonent_Bc_Trm.Value.item_number.ToString());
                            sw.WriteLine("               item_name = " + myComonent_Bc_Trm.Value.item_name.ToString());
                            sw.WriteLine("               type_of_soft_start = " + myComonent_Bc_Trm.Value.type_of_soft_start.ToString());
                            sw.WriteLine("               soft_time_interval = " + myComonent_Bc_Trm.Value.soft_time_interval.ToString());
                            sw.WriteLine("               reference_value = " + myComonent_Bc_Trm.Value.reference_value.ToString());
                            sw.WriteLine("               type_of_time_interpolation = " + myComonent_Bc_Trm.Value.type_of_time_interpolation.ToString());
                            sw.WriteLine("               type_of_space_interpolation = " + myComonent_Bc_Trm.Value.type_of_space_interpolation.ToString());
                            sw.WriteLine("            EndSect  // " + myComonent_Bc_Trm.Key.ToString());
                            sw.WriteLine("");
                        }
                    }
                    sw.WriteLine("         EndSect  // " + myCode_Bc_Trm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // BOUNDARY_CONDITIONS");
                myBoundary_Conditions_Trm = null;
                sw.WriteLine("");
                M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS myOutputs_Trm = myTransport_Module.outputs;
                sw.WriteLine("      [OUTPUTS]");
                sw.WriteLine("         Touched = " + myOutputs_Trm.Touched.ToString());
                sw.WriteLine("         MzSEPfsListItemCount = " + myOutputs_Trm.MzSEPfsListItemCount.ToString());
                sw.WriteLine("         number_of_outputs = " + myOutputs_Trm.number_of_outputs.ToString());
                foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT> myOutput_Trm in myOutputs_Trm.output)
                {
                    sw.WriteLine("         [" + myOutput_Trm.Key.ToString() + "]");
                    sw.WriteLine("            Touched = " + myOutput_Trm.Value.Touched.ToString());
                    sw.WriteLine("            include = " + myOutput_Trm.Value.include.ToString());
                    sw.WriteLine("            title = " + myOutput_Trm.Value.title.ToString());
                    sw.WriteLine("            file_name = " + myOutput_Trm.Value.file_name.ToString());
                    sw.WriteLine("            type = " + myOutput_Trm.Value.type.ToString());
                    sw.WriteLine("            format = " + myOutput_Trm.Value.format.ToString());
                    sw.WriteLine("            flood_and_dry = " + myOutput_Trm.Value.flood_and_dry.ToString());
                    sw.WriteLine("            coordinate_type = " + myOutput_Trm.Value.coordinate_type.ToString());
                    sw.WriteLine("            zone = " + myOutput_Trm.Value.zone.ToString());
                    sw.WriteLine("            input_file_name = " + myOutput_Trm.Value.input_file_name.ToString());
                    sw.WriteLine("            input_format = " + myOutput_Trm.Value.input_format.ToString());
                    sw.WriteLine("            interpolation_type = " + myOutput_Trm.Value.interpolation_type.ToString());
                    sw.WriteLine("            first_time_step = " + myOutput_Trm.Value.first_time_step.ToString());
                    sw.WriteLine("            last_time_step = " + myOutput_Trm.Value.last_time_step.ToString());
                    sw.WriteLine("            time_step_frequency = " + myOutput_Trm.Value.time_step_frequency.ToString());
                    sw.WriteLine("            number_of_points = " + myOutput_Trm.Value.number_of_points.ToString());
                    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT.POINT> myPoint_Out_Trm in myOutput_Trm.Value.point)
                    {
                        sw.WriteLine("            [" + myPoint_Out_Trm.Key.ToString() + "]");
                        sw.WriteLine("               name = " + myPoint_Out_Trm.Value.name.ToString());
                        sw.WriteLine("               x = " + myPoint_Out_Trm.Value.x.ToString());
                        sw.WriteLine("               y = " + myPoint_Out_Trm.Value.y.ToString());
                        sw.WriteLine("               z = " + myPoint_Out_Trm.Value.z.ToString());
                        sw.WriteLine("               layer = " + myPoint_Out_Trm.Value.layer.ToString());
                        sw.WriteLine("            EndSect  // " + myPoint_Out_Trm.Key.ToString());
                        sw.WriteLine("");
                    }
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT.LINE myLine_Out_Trm = myOutput_Trm.Value.line;
                    sw.WriteLine("            [LINE]");
                    sw.WriteLine("               npoints = " + myLine_Out_Trm.npoints.ToString());
                    sw.WriteLine("               x_first = " + myLine_Out_Trm.x_first.ToString());
                    sw.WriteLine("               y_first = " + myLine_Out_Trm.y_first.ToString());
                    sw.WriteLine("               z_first = " + myLine_Out_Trm.z_first.ToString());
                    sw.WriteLine("               x_last = " + myLine_Out_Trm.x_last.ToString());
                    sw.WriteLine("               y_last = " + myLine_Out_Trm.y_last.ToString());
                    sw.WriteLine("               z_last = " + myLine_Out_Trm.z_last.ToString());
                    sw.WriteLine("            EndSect  // LINE");
                    myLine_Out_Trm = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT.AREA myArea_Out_Trm = myOutput_Trm.Value.area;
                    sw.WriteLine("            [AREA]");
                    sw.WriteLine("               number_of_points = " + myArea_Out_Trm.number_of_points.ToString());
                    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT.AREA.POINT> myPoint_Area_Out_Trm in myArea_Out_Trm.point)
                    {
                        sw.WriteLine("               [" + myPoint_Area_Out_Trm.Key.ToString() + "]");
                        sw.WriteLine("                  x = " + myPoint_Area_Out_Trm.Value.x.ToString());
                        sw.WriteLine("                  y = " + myPoint_Area_Out_Trm.Value.y.ToString());
                        sw.WriteLine("               EndSect  // " + myPoint_Area_Out_Trm.Key.ToString());
                        sw.WriteLine("");
                    }
                    sw.WriteLine("               layer_min = " + myArea_Out_Trm.layer_min.ToString());
                    sw.WriteLine("               layer_max = " + myArea_Out_Trm.layer_max.ToString());
                    sw.WriteLine("            EndSect  // AREA");
                    myArea_Out_Trm = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT.PARAMETERS_2D myParameters_2d_Trm = myOutput_Trm.Value.parameters_2d;
                    sw.WriteLine("            [PARAMETERS_2D]");
                    sw.WriteLine("               Touched = " + myParameters_2d_Trm.Touched.ToString());
                    sw.WriteLine("               COMPONENT_1 = " + myParameters_2d_Trm.COMPONENT_1.ToString());
                    sw.WriteLine("               U_VELOCITY = " + myParameters_2d_Trm.U_VELOCITY.ToString());
                    sw.WriteLine("               V_VELOCITY = " + myParameters_2d_Trm.V_VELOCITY.ToString());
                    sw.WriteLine("               CFL_NUMBER = " + myParameters_2d_Trm.CFL_NUMBER.ToString());
                    sw.WriteLine("            EndSect  // PARAMETERS_2D");
                    myParameters_2d_Trm = null;
                    sw.WriteLine("");

                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT.PARAMETERS_3D myParameters_3d_Trm = myOutput_Trm.Value.parameters_3d;
                    sw.WriteLine("            [PARAMETERS_3D]");
                    sw.WriteLine("               Touched = " + myParameters_3d_Trm.Touched.ToString());
                    sw.WriteLine("               COMPONENT_1 = " + myParameters_3d_Trm.COMPONENT_1.ToString());
                    sw.WriteLine("               U_VELOCITY = " + myParameters_3d_Trm.U_VELOCITY.ToString());
                    sw.WriteLine("               V_VELOCITY = " + myParameters_3d_Trm.V_VELOCITY.ToString());
                    sw.WriteLine("               W_VELOCITY = " + myParameters_3d_Trm.W_VELOCITY.ToString());
                    sw.WriteLine("               CFL_NUMBER = " + myParameters_3d_Trm.CFL_NUMBER.ToString());
                    sw.WriteLine("            EndSect  // PARAMETERS_3D");
                    myParameters_3d_Trm = null;
                    sw.WriteLine("");

                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT.DISCHARGE myDischarge_Trm = myOutput_Trm.Value.discharge;
                    sw.WriteLine("            [DISCHARGE]");
                    sw.WriteLine("               Touched = " + myDischarge_Trm.Touched.ToString());
                    sw.WriteLine("               COMPONENT_1 = " + myDischarge_Trm.COMPONENT_1.ToString());
                    sw.WriteLine("            EndSect  // DISCHARGE");
                    myDischarge_Trm = null;
                    sw.WriteLine("");
                    M21_3FMService.FemEngineHD.TRANSPORT_MODULE.OUTPUTS.OUTPUT.MASSBUDGET myMassbudget_Trm = myOutput_Trm.Value.massbudget;
                    sw.WriteLine("            [MASSBUDGET]");
                    sw.WriteLine("               Touched = " + myMassbudget_Trm.Touched.ToString());
                    sw.WriteLine("               COMPONENT_1 = " + myMassbudget_Trm.COMPONENT_1.ToString());
                    sw.WriteLine("            EndSect  // MASSBUDGET");
                    myMassbudget_Trm = null;
                    sw.WriteLine("");
                    sw.WriteLine("         EndSect  // " + myOutput_Trm.Key.ToString());
                    sw.WriteLine("");
                }
                sw.WriteLine("      EndSect  // OUTPUTS");
                myOutputs_Trm = null;
                sw.WriteLine("");
                sw.WriteLine("   EndSect  // TRANSPORT_MODULE");
                myTransport_Module = null;
                sw.WriteLine("");
                ////////////////////ECOLAB_MODULE///////////////////////////
                sw.WriteLine(myFemEnginHD.ecolab_module);
                sw.WriteLine(myFemEnginHD.mud_transport_module);
                sw.WriteLine(myFemEnginHD.sand_transport_module);
                sw.WriteLine(myFemEnginHD.particle_tracking_module);

                sw.WriteLine("EndSect  // FemEngineHD");
                sw.WriteLine("");
                sw.WriteLine("[SYSTEM]");
                sw.WriteLine("   ResultRootFolder = " + m21_3fmInput.system.ResultRootFolder.ToString());
                sw.WriteLine("   UseCustomResultFolder = " + ((bool)m21_3fmInput.system.UseCustomResultFolder).ToString().ToLower());
                sw.WriteLine("   CustomResultFolder = " + m21_3fmInput.system.CustomResultFolder.ToString());
                sw.WriteLine("EndSect  // SYSTEM");
                myFemEnginHD = null;
                sw.WriteLine("");
                sw.Flush();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotWriteFile_Error_, fi.FullName, ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotWriteFile_Error_", fi.FullName, ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
            }
        }
        private string expandList(string lineStart, List<float> list)
        {
            StringBuilder mySb = new StringBuilder();
            mySb.Append(lineStart);
            for (int i = 0; i <= list.Count - 1; i++)
            {
                mySb.Append(list[i].ToString().Trim() + ", ");
            }
            mySb.Remove(mySb.Length - 2, 1);
            return mySb.ToString().TrimEnd();
        }
        private string expandList(string lineStart, List<double> list)
        {
            StringBuilder mySb = new StringBuilder();
            mySb.Append(lineStart);
            for (int i = 0; i <= list.Count - 1; i++)
            {
                mySb.Append(list[i].ToString().Trim() + ", ");
            }
            mySb.Remove(mySb.Length - 2, 1);
            return mySb.ToString().TrimEnd();
        }
        private string expandList(string lineStart, List<string> list)
        {
            StringBuilder mySb = new StringBuilder();
            mySb.Append(lineStart);
            for (int i = 0; i <= list.Count - 1; i++)
            {
                mySb.Append(list[i].ToString().Trim() + ", ");
            }
            mySb.Remove(mySb.Length - 2, 1);
            return mySb.ToString().TrimEnd();
        }
        private string expandList(string lineStart, List<int> list)
        {
            StringBuilder mySb = new StringBuilder();
            mySb.Append(lineStart);
            for (int i = 0; i <= list.Count - 1; i++)
            {
                mySb.Append(list[i].ToString().Trim() + ", ");
            }
            mySb.Remove(mySb.Length - 2, 1);
            return mySb.ToString().TrimEnd();
        }
    }

}
