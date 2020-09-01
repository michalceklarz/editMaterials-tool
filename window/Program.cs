using System;
using Xbim.Ifc;
using Xbim.IO;
using Xbim.Ifc4.MaterialResource;
using Xbim.Ifc4.Interfaces;
using Xbim.Ifc4.ProductExtension;
using Newtonsoft.Json.Linq;
using Xbim.Ifc4.SharedBldgElements;
using Xbim.Ifc4.RepresentationResource;
using Xbim.Ifc4.ArchitectureDomain;
using IfcRelDefinesByType = Xbim.Ifc4.Kernel.IfcRelDefinesByType;
using Xbim.Ifc4.GeometricModelResource;
using Xbim.Ifc4.TopologyResource;
using Xbim.Ifc4.GeometryResource;
using Xbim.Ifc4.GeometricConstraintResource;
using System.Collections.Generic;
using Xbim.Ifc4.Kernel;
using Xbim.Ifc2x3.Kernel;
using IfcPropertySet = Xbim.Ifc2x3.Kernel.IfcPropertySet;
using Xbim.Ifc4.PropertyResource;
using Xbim.Ifc2x3.PropertyResource;
using IfcPropertySingleValue = Xbim.Ifc4.PropertyResource.IfcPropertySingleValue;
using IfcRelDefinesByProperties = Xbim.Ifc4.Kernel.IfcRelDefinesByProperties;
using Xbim.Ifc4.MeasureResource;
using Xbim.Ifc2x3.MeasureResource;
using IfcSimpleValue = Xbim.Ifc4.MeasureResource.IfcSimpleValue;
using IfcMeasureValue = Xbim.Ifc2x3.MeasureResource.IfcMeasureValue;
using IfcCountMeasure = Xbim.Ifc4.MeasureResource.IfcCountMeasure;
using IfcSIUnit = Xbim.Ifc4.MeasureResource.IfcSIUnit;
using IfcUnitEnum = Xbim.Ifc4.Interfaces.IfcUnitEnum;
using IfcPositiveLengthMeasure = Xbim.Ifc4.MeasureResource.IfcPositiveLengthMeasure;
using IfcLabel = Xbim.Ifc4.MeasureResource.IfcLabel;
using Xbim.Ifc2x3.SharedBldgElements;
using IfcWindowPanelProperties = Xbim.Ifc4.ArchitectureDomain.IfcWindowPanelProperties;
using IfcWindowLiningProperties = Xbim.Ifc4.ArchitectureDomain.IfcWindowLiningProperties;
using IfcWindow = Xbim.Ifc4.SharedBldgElements.IfcWindow;
using IfcPositiveRatioMeasure = Xbim.Ifc4.MeasureResource.IfcPositiveRatioMeasure;
using IfcWindowStyle = Xbim.Ifc4.ArchitectureDomain.IfcWindowStyle;
using IfcObject = Xbim.Ifc4.Kernel.IfcObject;
using System.Linq;
using IfcPropertyDefinition = Xbim.Ifc4.Kernel.IfcPropertyDefinition;
using IfcProperty = Xbim.Ifc4.PropertyResource.IfcProperty;

namespace window
{
    class Program
    {
        static void Main(string[] args)
        {

            //Geting path to the IFC file
            string[] lines = System.IO.File.ReadAllLines(@AppDomain.CurrentDomain.BaseDirectory + "path.txt");
            string fileName = lines[0];
            Console.WriteLine("Editing: " + fileName);

            //Geting material information
            string data = System.IO.File.ReadAllText(@AppDomain.CurrentDomain.BaseDirectory + "layers_data_write.json");

            dynamic json = JArray.Parse(data);

            //Editor credentials and access mode
            var editor = new XbimEditorCredentials
            {
                //Maybe add sth later

            };
            Console.WriteLine("Editor: ");

            XbimDBAccess accessmode = XbimDBAccess.Exclusive;

            using (var model = IfcStore.Open(fileName, editor, null, null, accessmode))
            {

                using (var txn = model.BeginTransaction("Widnows modification"))
                {
                    foreach (dynamic element in json)
                    {

                        //Read data:

                        string id = element.id;
                        Console.WriteLine(element.id);
                        string glazing_material = element.glazing_material;
                        string lining_material = element.lining_material;

                        double lining_thickness = element.lining_thickness;
                        int lining_partitioning = element.lining_partitioning;
                        double framing_thickness = element.framing_thickness;
                        double framing_mulion_thickness = element.framing_mulion_thickness;
                        double framing_transom_thickness = element.framing_transom_thickness;

                        //Units consideration:
                        var units = model.Instances.FirstOrDefault<IfcSIUnit>(u => u.UnitType == IfcUnitEnum.LENGTHUNIT);
                        Console.WriteLine(units.FullName);
                        
                        lining_thickness = lining_thickness / 1000;
                        framing_thickness = framing_thickness / 1000;
                        framing_mulion_thickness = framing_mulion_thickness / 1000;
                        framing_transom_thickness = framing_transom_thickness / 1000;
                        

                        // Get window for material editing 
                        var window = model.Instances.FirstOrDefault<IfcWindow>(w => w.GlobalId == id);

                        // Materials

                        var ifcWindowType = GetIfcWindowType(window);

                        var ifcMaterialConstituentSet = GetIfcMaterialConstituentSet(ifcWindowType);
                        ClearMaterialConstituents(ifcMaterialConstituentSet);

                        var ifcMaterial_lining = model.Instances.New<IfcMaterial>();
                        ifcMaterial_lining.Name = lining_material;

                        var ifcMaterialConstituent_lining = model.Instances.New<IfcMaterialConstituent>();
                        ifcMaterialConstituent_lining.Material = ifcMaterial_lining;
                        ifcMaterialConstituentSet.MaterialConstituents.Add(ifcMaterialConstituent_lining);
                        ifcMaterialConstituent_lining.Name = "Lining";


                        var ifcMaterial_glazing = model.Instances.New<IfcMaterial>();
                        ifcMaterial_glazing.Name = glazing_material;

                        var ifcMaterialConstituent_glazing = model.Instances.New<IfcMaterialConstituent>();
                        ifcMaterialConstituent_glazing.Material = ifcMaterial_glazing;
                        ifcMaterialConstituentSet.MaterialConstituents.Add(ifcMaterialConstituent_glazing);
                        ifcMaterialConstituent_glazing.Name = "Glazing";

                        


                        //IfcWindowLiningProperties



                        var ifcWindowLiningProperties = GetIfcWindowLiningProperties(ifcWindowType);

                        ifcWindowLiningProperties.LiningThickness = lining_thickness;
                        //ifcWindowLiningProperties.MullionThickness = framing_mulion_thickness;
                        ifcWindowLiningProperties.TransomThickness = framing_transom_thickness;




                        





                        ifcWindowType.PartitioningType = (IfcWindowTypePartitioningEnum)lining_partitioning;


                        //IfcWindowPanelProperties 

                       ClearPanelProperties(ifcWindowType);

                        if (lining_partitioning == 0)   NewWindowPanelProperty(ifcWindowType,"MIDDLE",framing_thickness  );
                        if (lining_partitioning == 1) { NewWindowPanelProperty(ifcWindowType, "BOTTOM", framing_thickness); NewWindowPanelProperty(ifcWindowType, "TOP", framing_thickness);}
                        if (lining_partitioning == 2) { NewWindowPanelProperty(ifcWindowType, "LEFT", framing_thickness  ); NewWindowPanelProperty(ifcWindowType, "RIGHT", framing_thickness); }
                        if (lining_partitioning == 3) { NewWindowPanelProperty(ifcWindowType, "LEFT", framing_thickness  ); NewWindowPanelProperty(ifcWindowType, "MIDDLE", framing_thickness); NewWindowPanelProperty(ifcWindowType, "RIGHT", framing_thickness); }
                        if (lining_partitioning == 8) { NewWindowPanelProperty(ifcWindowType, "BOTTOM", framing_thickness); NewWindowPanelProperty(ifcWindowType, "MIDDLE", framing_thickness); NewWindowPanelProperty(ifcWindowType, "TOP", framing_thickness); }
                        


                        //IfcWindow Pset_WindowCommon

                        var GlazingAreaFraction = window.GetPropertySingleValue("Pset_WindowCommon", "GlazingAreaFraction");

                        if (GlazingAreaFraction == null)
                        {
                            var ifcPropertySet = model.Instances.New<Xbim.Ifc4.Kernel.IfcPropertySet>();
                            ifcPropertySet.Name = "Pset_WindowCommon";
                            GlazingAreaFraction = model.Instances.New<IfcPropertySingleValue>();
                            GlazingAreaFraction.Name = "GlazingAreaFraction";
                            ifcPropertySet.HasProperties.Add((IfcProperty)GlazingAreaFraction);
                        }


                        GlazingAreaFraction.NominalValue = CalculateGlazingAreaFraction(window,
                                                                                        lining_partitioning,
                                                                                        framing_transom_thickness,
                                                                                        framing_mulion_thickness,
                                                                                        framing_thickness,
                                                                                        lining_thickness);



                        //Manual IFC sample model
                        ifcWindowLiningProperties.LiningDepth = 0.02;

                        ifcWindowLiningProperties.FirstTransomOffset = window.OverallHeight / 2;

                        var ifcMaterialConstituent_framing = model.Instances.New<IfcMaterialConstituent>();
                        ifcMaterialConstituent_framing.Material = ifcMaterial_lining;
                        ifcMaterialConstituentSet.MaterialConstituents.Add(ifcMaterialConstituent_framing);
                        ifcMaterialConstituent_framing.Name = "Framing";



                        var Pset_DoorWindowGlazingType = model.Instances.New<Xbim.Ifc4.Kernel.IfcPropertySet>();
                        Pset_DoorWindowGlazingType.Name = "	Pset_DoorWindowGlazingType"; 


                        var GlassLayers = model.Instances.New<IfcPropertySingleValue>();
                        GlassLayers.Name = "GlassLayers";
                        
                        Pset_DoorWindowGlazingType.HasProperties.Add(GlassLayers);

                        var GlassThickness1 = model.Instances.New<IfcPropertySingleValue>();
                        GlassThickness1.Name = "GlassThickness1";
                        Pset_DoorWindowGlazingType.HasProperties.Add(GlassThickness1);
                        IfcPositiveLengthMeasure ifcPositiveLengthMeasure0 = 0.002;
                        GlassThickness1.NominalValue = ifcPositiveLengthMeasure0;

                        var GlassThickness2 = model.Instances.New<IfcPropertySingleValue>();
                        GlassThickness2.Name = "GlassThickness2";
                        Pset_DoorWindowGlazingType.HasProperties.Add(GlassThickness2);
                        IfcPositiveLengthMeasure ifcPositiveLengthMeasure1 = 0.002;
                        GlassThickness2.NominalValue = ifcPositiveLengthMeasure1;

                        var FillGas = model.Instances.New<IfcPropertySingleValue>();
                        FillGas.Name = "FillGas";
                        IfcLabel gass = "Argon";
                        FillGas.NominalValue = gass;
                        Pset_DoorWindowGlazingType.HasProperties.Add(FillGas);

                        var glass_layers = model.Instances.New<IfcPropertySingleValue>();
                        glass_layers.Name = "GlassLayers";
                        IfcCountMeasure two = 2;
                        GlassLayers.NominalValue = two;

                    }





                    txn.Commit();
                    model.SaveAs(fileName);
                }



                IfcMaterialConstituentSet GetIfcMaterialConstituentSet(IfcWindowType type)
                {

                    IfcMaterialConstituentSet ifcMaterialConstituentSet;
                    var ifcRelAssociateMaterial = model.Instances.FirstOrDefault<IfcRelAssociatesMaterial>(r=>r.RelatedObjects.First() == type);
                    if (ifcRelAssociateMaterial == null) 
                    {   
                        Console.WriteLine("New ifcRelAssociateMaterial and ifcMaterialConstituentSet");
                        ifcRelAssociateMaterial = model.Instances.New<IfcRelAssociatesMaterial>();
                        ifcMaterialConstituentSet =  model.Instances.New<IfcMaterialConstituentSet>();
                        RelatesMaterialToWindowType(type, ifcMaterialConstituentSet, ifcRelAssociateMaterial);
                        return ifcMaterialConstituentSet;
                    }
                    else
                    {
                        ifcMaterialConstituentSet = model.Instances.FirstOrDefault<IfcMaterialConstituentSet>(m => m == ifcRelAssociateMaterial.RelatingMaterial);
                        if (ifcMaterialConstituentSet == null)
                        {
                            Console.WriteLine("New ifcMaterialConstituentSet");
                            ifcMaterialConstituentSet = model.Instances.New<IfcMaterialConstituentSet>();
                            RelatesMaterialToWindowType(type, ifcMaterialConstituentSet, ifcRelAssociateMaterial);
                            return ifcMaterialConstituentSet;
                        }
                        else
                        {
                            Console.WriteLine("Reused ifcMaterialConstituentSet"); 
                            return ifcMaterialConstituentSet;
                        }
                    }

                }

                IfcWindowType GetIfcWindowType(IfcObject element)
                {
                    IfcWindowType ifcWindowType;
                    var ifcRelDefinesByType = model.Instances.FirstOrDefault<IfcRelDefinesByType>(r => r.RelatedObjects.FirstOrDefault() == element);
                    if (ifcRelDefinesByType == null) 
                    { 
                        ifcWindowType = model.Instances.New<IfcWindowType>();
                        RelatesWindowToWindowType((IfcWindow)element, ifcWindowType);
                        return ifcWindowType;
                    }
                        
                    else
                    {
                        ifcWindowType = model.Instances.FirstOrDefault<IfcWindowType>(w => w == ifcRelDefinesByType.RelatingType);

                        if (ifcWindowType == null) 
                        {
                            ifcWindowType =  model.Instances.New<IfcWindowType>();
                            ifcRelDefinesByType.RelatingType = ifcWindowType;
                            return ifcWindowType;
                        }

                        else return ifcWindowType;
                    }

                }

                IfcRelDefinesByType GetIfcRelDefinesByType(IfcObject element)
                {
                    IfcRelDefinesByType ifcRelDefinesByType;
                     ifcRelDefinesByType = model.Instances.FirstOrDefault<IfcRelDefinesByType>(r => r.RelatedObjects == element);
                    if (ifcRelDefinesByType == null) 
                    { 
                        ifcRelDefinesByType = model.Instances.New<IfcRelDefinesByType>();
                        ifcRelDefinesByType.RelatedObjects.Add(element);
                        return ifcRelDefinesByType;
                    }
                    else return ifcRelDefinesByType;
                }

                void RelatesMaterialToWindowType(IfcWindowType type, IfcMaterialConstituentSet material, IfcRelAssociatesMaterial associates)
                {
                    associates.RelatingMaterial = material;
                    associates.RelatedObjects.Add(type);
                }

                void RelatesWindowToWindowType(IfcWindow window, IfcWindowType type)
                    {
                        var ifcRelDefinesByType = GetIfcRelDefinesByType(window);
                        ifcRelDefinesByType.RelatingType = type;
                    }


                Xbim.Ifc4.Kernel.IfcPropertySet GetPset_DoorWindowGlazingType(IfcObject element)
                {
                    Xbim.Ifc4.Kernel.IfcPropertySet Pset;
                    var ifcRelDefinesByProperties = model.Instances.FirstOrDefault<IfcRelDefinesByProperties>(r => r.RelatedObjects == element);
                    if (ifcRelDefinesByProperties == null) 
                    { 
                        Pset = model.Instances.New<Xbim.Ifc4.Kernel.IfcPropertySet>();
                        Pset.Name = "Pset_DoorWindowGlazingType";
                        ifcRelDefinesByProperties = model.Instances.New<IfcRelDefinesByProperties>();
                        ifcRelDefinesByProperties.RelatedObjects.Add(element);
                        ifcRelDefinesByProperties.RelatingPropertyDefinition = Pset;
                        return Pset; 
                    }
                    else Pset = model.Instances.FirstOrDefault<Xbim.Ifc4.Kernel.IfcPropertySet>(p=>p.Name == "Pset_DoorWindowGlazingType");
                    if (Pset == null) 
                    { 
                        Pset = model.Instances.New<Xbim.Ifc4.Kernel.IfcPropertySet>(); 
                        Pset.Name = "Pset_DoorWindowGlazingType";
                        ifcRelDefinesByProperties.RelatingPropertyDefinition = Pset;
                        return Pset; 
                    }
                    else return Pset;
                }

                
                

                void ClearMaterialConstituents(IfcMaterialConstituentSet set)
                {
                    int number = set.MaterialConstituents.Count;
                    for (int i =0;i< number; i++)
                    {
                        
                        model.Delete(set.MaterialConstituents.First().Material);
                        model.Delete(set.MaterialConstituents.First());
                        
                    }

                }

                void ClearPset_DoorWindowGlazingType(Xbim.Ifc4.Kernel.IfcPropertySet Pset)
                {
                    for (int i = 0; i < Pset.HasProperties.Count; i++)
                    {
                        model.Delete(Pset.HasProperties.First());
                    }
                }

                IfcWindowLiningProperties GetIfcWindowLiningProperties(IfcWindowType type)
                {
                    IfcWindowLiningProperties ifcWindowLiningProperties;
                    ifcWindowLiningProperties = model.Instances.FirstOrDefault<IfcWindowLiningProperties>(w=>w.DefinesType.FirstOrDefault()== type);
                    if (ifcWindowLiningProperties == null)
                    {
                        Console.WriteLine("Create new IfcWindowLiningProperties");
                        ifcWindowLiningProperties = model.Instances.New<IfcWindowLiningProperties>();
                        type.HasPropertySets.Add(ifcWindowLiningProperties);
                        return ifcWindowLiningProperties;
                    }
                    else
                    {
                        Console.WriteLine("Reuse IfcWindowLiningProperties");
                        return ifcWindowLiningProperties;
                    }
                }

                void NewWindowPanelProperty(IfcWindowType type, string position,double frame_thickness)
                {
                   var ifcWindowPanelProperties = model.Instances.New<IfcWindowPanelProperties>(); 
                   type.HasPropertySets.Add(ifcWindowPanelProperties);
                   ifcWindowPanelProperties.FrameThickness = frame_thickness;
                    
                    if (position == "LEFT") ifcWindowPanelProperties.PanelPosition = Xbim.Ifc4.Interfaces.IfcWindowPanelPositionEnum.LEFT;
                    if (position == "MIDDLE") ifcWindowPanelProperties.PanelPosition = Xbim.Ifc4.Interfaces.IfcWindowPanelPositionEnum.MIDDLE;
                    if (position == "RIGHT") ifcWindowPanelProperties.PanelPosition = Xbim.Ifc4.Interfaces.IfcWindowPanelPositionEnum.RIGHT;
                    if (position == "BOTTOM") ifcWindowPanelProperties.PanelPosition = Xbim.Ifc4.Interfaces.IfcWindowPanelPositionEnum.BOTTOM;
                    if (position == "TOP") ifcWindowPanelProperties.PanelPosition = Xbim.Ifc4.Interfaces.IfcWindowPanelPositionEnum.TOP;


                    //Manual sample IFC window model
                    ifcWindowPanelProperties.FrameDepth = 0.05;
                    //


                    Console.WriteLine(ifcWindowPanelProperties+" "+ ifcWindowPanelProperties.PanelPosition);
                }

                void ClearPanelProperties(IfcWindowType type)
                {
                    var allproperties = model.Instances.OfType<IfcWindowPanelProperties>();
                    Console.WriteLine(model.Instances.OfType<IfcWindowPanelProperties>().Count());

                    for (int i = 0; i <  model.Instances.OfType<IfcWindowPanelProperties>().Count(); i++)
                    {
                        if (allproperties.ElementAt(i).DefinesType.FirstOrDefault() == type) { Console.WriteLine(allproperties.ElementAt(i)); model.Delete(allproperties.ElementAt(i)); i--;  }
                    }

                }

                IfcPositiveRatioMeasure CalculateGlazingAreaFraction(IfcWindow window, double lining_partitioning,  double framing_transom_thickness, double framing_mulion_thickness, double framing_thickness, double lining_thickness)
                {
                    double height = window.OverallHeight.Value;
                    double width = window.OverallWidth.Value;

                    double non_glazing_area = 0;
                    double non_glazing_area_single_panel = 2 * (lining_thickness + framing_thickness) * (height + width - 2 * (lining_thickness + framing_thickness));
                    switch (lining_partitioning)
                    {
                        case 0:     
                            non_glazing_area = non_glazing_area_single_panel;
                            break;
                        case 1:
                            non_glazing_area = non_glazing_area_single_panel + (height - 2 * (lining_thickness + framing_thickness)) * (framing_transom_thickness + 2 * framing_thickness);
                            break;
                        case 2:
                            non_glazing_area = non_glazing_area_single_panel + (width - 2 * (lining_thickness + framing_thickness)) * (framing_mulion_thickness + 2 * framing_thickness);
                            break;
                        case 3:
                            non_glazing_area = non_glazing_area_single_panel + 2 * (height - 2 * (lining_thickness + framing_thickness)) * (framing_transom_thickness + 2 * framing_thickness);
                            break;
                        case 8:
                            non_glazing_area = non_glazing_area_single_panel + 2 * (width - 2 * (lining_thickness + framing_thickness)) * (framing_mulion_thickness + 2 * framing_thickness);
                            break;
                        default:
                            Console.WriteLine("Default case");
                            break;

                    }

                    double GlazingAreaFraction_double = (height * width - non_glazing_area) / (height * width);
                    Console.WriteLine(GlazingAreaFraction_double);
                    IfcPositiveRatioMeasure GlazingAreaFraction_value = GlazingAreaFraction_double;
                    
                    return GlazingAreaFraction_value;
                }
            }
        }
    }
}
