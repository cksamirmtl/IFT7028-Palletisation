using System;
using System.Activities.Expressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*Add from NuGet:
LinkToExcel v1.11.0
SimSharp v3.4.1*/

namespace IFT7028_Palletisation
{
    internal class Program
    {
        //******Paramètres******//
        private static string xlsxInputFilePath = @"C:\Local\E2022\Conception et simulation des systèmes intelligents pour l'industrie 4.0 - (IFT-7028)\Projet\20220603\DataSet.xlsx";
        private static string panneauxWorksheet = @"C-3312";
        private static string palettesWorksheet = @"C-3312_palette";
        //private static int nbPalettesSurPlancher = 5;
        //Inter arrivé des paneaux entre 3 et 5 minutes
        private static readonly SimSharp.UniformTime PanneauxInterArrivalTime = SimSharp.Distributions.UNIF(new TimeSpan(0, 0, 0), new TimeSpan(0, 0, 1));
        //Temps traitement panneau 3 et 5 minutes
        private static readonly SimSharp.UniformTime PanneauxTraitementTime = SimSharp.Distributions.UNIF(new TimeSpan(0, 0, 0), new TimeSpan(0, 0, 1));
        //Temps manutention lorsqu'une palette est terminée (entre 5 et 10 minutes)
        private static readonly SimSharp.UniformTime PaletteHandlingTime = SimSharp.Distributions.UNIF(new TimeSpan(0, 0, 0), new TimeSpan(0, 0, 1));
        //Simulation time in hours
        private static readonly TimeSpan SimTime = TimeSpan.FromHours(10000);
        //Initialiser avec des zéros l'ensemble de palettes simultannées
        //static List<PaletteStruct> palettes_encours1 = new List<PaletteStruct>()
        static List<object> palettes_encours1 = new List<object>()
        {
            new PaletteStruct() {num_palette = 0, quantity = 0},
            new PaletteStruct() {num_palette = 0, quantity = 0},
            new PaletteStruct() {num_palette = 0, quantity = 0},
            new PaletteStruct() {num_palette = 0, quantity = 0},
            new PaletteStruct() {num_palette = 0, quantity = 0},
        };
        //******Paramètres******//


        IEnumerable<object> palettes_encours2 = (IEnumerable<object>)palettes_encours1;

        public struct PanneauStruct
        {
            public string No;
            public string Id;
            public string Type;
            public int Secteur;
            public int Num_palette;
        }

        public struct PaletteStruct
        {
            public int num_palette;
            public int quantity;
        }

        //PaletteStruct[] palettes_encours = new PaletteStruct[nbPalettesSurPlancher];

        DateTime simulationStartTime;

        delegate Func<object, bool> SelectorDelegate(object palletisationService);
        Func<object, bool> SelectorDelegateImp = delegate (object palletisationService)
        {
            if (palettes_encours1.Exists(x => ((PaletteStruct)x).num_palette == ((MyObject)palletisationService).panneau.Num_palette))
                return true;
            //if (palettes_encours1.Count < nbPalettesSurPlancher)
            //    return true;
            foreach (PaletteStruct palette_encours in palettes_encours1)
            {
                //si une palette est pleine, ou une palette fictive: on la remplace par la pelette qu'on a besoin pour notre panneau
                if ((palette_encours.num_palette == 0) || (all_palettes.First(palette => palette.num_palette == palette_encours.num_palette).quantity == palette_encours.quantity))
                {
                    return true;
                }
            }
            return false;
        };

        public struct MyObject
        {
            public PanneauStruct panneau;
            public SimSharp.ResourcePool palletisationService;
        }

        static IQueryable<PaletteStruct> all_palettes;
        static IQueryable<PanneauStruct> all_panneau;

        private static object myLock = new object();
        static void Main(string[] args)
        {
            Console.WriteLine("Enter to start..");
            Console.ReadLine();
            new Program().Simulate();
            Console.WriteLine("Done.");
            Console.ReadLine();
        }

        public void Simulate()
        {
            int RandomSeed = 77;
            var env = new SimSharp.Simulation(RandomSeed);
            //palettes_encours1.Where(p => ((PaletteStruct)p).num_palette == 2);
            //var palletisationService = new SimSharp.Resource(env, nbPalettesSurPlancher)
            //var palletisationService = new SimSharp.ResourcePool(env, (IEnumerable<object>)(System.Collections.IEnumerable)palettes_encours)
            var palletisationService = new SimSharp.ResourcePool(env, (IEnumerable<object>)palettes_encours2)
            {
                //des metriques
                QueueLength = new SimSharp.TimeSeriesMonitor(env, name: "Queue", collect: true),
                //
            };

            var excelFile = new LinqToExcel.ExcelQueryFactory(xlsxInputFilePath);
            all_panneau =
                from row in excelFile.Worksheet(panneauxWorksheet)
                let item = new PanneauStruct()
                {
                    No = row["No"].Cast<string>(),
                    Id = row["Id"].Cast<string>(),
                    Type = row["Type"].Cast<string>(),
                    Secteur = row["Secteur"].Cast<int>(),
                    Num_palette = row["Palette"].Cast<int>(),
                }
                where item.Num_palette > 0
                select item;

            all_palettes =
                from row in excelFile.Worksheet(palettesWorksheet)
                let item = new PaletteStruct()
                {
                    num_palette = row["Palette"].Cast<int>(),
                    quantity = row["qty"].Cast<int>(),
                }
                where item.num_palette > 0
                select item;

            env.Process(PanneauxGenerator(env, palletisationService));
            /***RUN***/
            simulationStartTime = env.Now;
            env.Run(SimTime);
            /*********/
        }

        private IEnumerable<SimSharp.Event> PanneauxGenerator(SimSharp.Simulation env, SimSharp.ResourcePool palletisationService)
        {
            foreach (PanneauStruct panneau in all_panneau)
            {
                //Console.WriteLine("PanneauNo {0}:", panneau.No);
                //yield return env.Timeout(PanneauxInterArrivalTime);
                yield return env.Timeout(new TimeSpan(0));
                env.Process(PanneauProcess(panneau, env, palletisationService));
            }
        }

        //Func<object, bool> selector = palettes_encours => ((IEnumerable<object>)(System.Collections.IEnumerable)palettes_encours).;
        //Func<object, bool> selector = x => ((IEnumerable<object>)(System.Collections.IEnumerable)x).;
        //Func<object, bool> selector = x => ((PaletteStruct)x).num_palette == 2;

        private IEnumerable<SimSharp.Event> PanneauProcess(PanneauStruct panneau, SimSharp.Simulation env, SimSharp.ResourcePool palletisationService)
        {
            DateTime arrivingTime = env.Now;
            env.Log("Processing Panneau: " + panneau.No.ToString());

            //Request access to palletisation service
            MyObject myObject = new MyObject();
            myObject.panneau = panneau;
            myObject.palletisationService = palletisationService;
            //SelectorDelegate selector = SelectorDelegateImp;
            using (var req = palletisationService.Request(new Func<object, bool>((o) => SelectorDelegateImp(myObject))))
            {
                //Request access to palletisation service
                Again:
                yield return req;

                //Checker si la palette pour notre panneau est disponible, sinon mettre le panneau dans le Tampon
                //Ce check doit être fait après l'autorisation d'accès au service, car la liste des palettes en cours peut changer

                //mettre à jour la liste palettes_encours
                lock (myLock)
                {
                    if (palettes_encours1.Exists(x => ((PaletteStruct)x).num_palette == panneau.Num_palette))
                    {
                        int i = palettes_encours1.FindIndex(x => ((PaletteStruct)x).num_palette == panneau.Num_palette);
                        PaletteStruct pl = ((PaletteStruct)palettes_encours1[i]);
                        pl.quantity++;
                        if (pl.quantity <= all_palettes.First(x => x.num_palette == pl.num_palette).quantity)
                            palettes_encours1[i] = pl;
                        else
                            goto Again;
                    }
                    else
                    {
                        foreach (PaletteStruct palette_encours in palettes_encours1.ToList())
                        {
                            //si une palette est pleine, ou une palette fictive: on la remplace par la pelette qu'on a besoin pour notre panneau
                            if ((palette_encours.num_palette == 0) || (all_palettes.First(palette => palette.num_palette == palette_encours.num_palette).quantity == palette_encours.quantity))
                            {
                                if (!(palette_encours.num_palette == 0)) 
                                {
                                    //yield return env.Timeout(PaletteHandlingTime);
                                }
                                //else break;
                                PaletteStruct pl = all_palettes.First(palette => palette.num_palette == panneau.Num_palette);
                                pl.quantity = 1;
                                int i = palettes_encours1.FindIndex(x => ((PaletteStruct)x).num_palette == palette_encours.num_palette);
                                if (i != -1)
                                    palettes_encours1[i] = pl;
                                else 
                                    goto Again;
                                break;
                            }
                        }
                    }
                }
                Console.Write("Etat palettes: ");
                foreach (PaletteStruct palette_encours in palettes_encours1)
                {
                    Console.Write(palette_encours.num_palette.ToString() + ": " + palette_encours.quantity.ToString() + "/");
                    int qtyTotal = 0;
                    try
                    {
                        qtyTotal = all_palettes.First(x => x.num_palette == palette_encours.num_palette).quantity;
                    }
                    catch (Exception ex) { }
                    Console.Write(qtyTotal.ToString() + " - ");
                }
                Console.WriteLine();
                // Traitement panneau
                //yield return env.Timeout(PanneauxTraitementTime);

            }
        }
    }
}
