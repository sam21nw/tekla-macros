// Generated by mihu
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Collections;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Model;

namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static Tekla.Technology.Akit.IScript akit;
        public static void Run(Tekla.Technology.Akit.IScript akit_in)
        {
            System.Runtime.Remoting.Lifetime.ClientSponsor sponsor = null;
            try
            {
                sponsor = new System.Runtime.Remoting.Lifetime.ClientSponsor();
                akit = akit_in;
                SelectThem();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
            finally
            {
                if (sponsor != null)
                {
                    sponsor.Close();
                }
            }
        }

        private static void SelectThem()
        {
            var list = new ArrayList();
            var ms = new Tekla.Structures.Model.UI.ModelObjectSelector();
            var s = ms.GetSelectedObjects().GetEnumerator();
            while (s.MoveNext())
            {
                var rmo = s.Current as ReferenceModelObject;
                if (rmo == null)
                {
                    continue;
                }

                var guid = string.Empty;
                rmo.GetUserProperty("Converted GUID", ref guid);
                if (!string.IsNullOrEmpty(guid))
                {
                    guid = guid.ToLower();
                    if (guid.StartsWith("id"))
                    {
                        guid = guid.Remove(0, 2);
                    }

                    try
                    {
                        var m = new Beam();
                        m.Identifier.GUID = new Guid(guid);
                        m.Select();
                        list.Add(m);
                    }
                    catch { }
                }
            }

            ms.Select(new ArrayList());
            ms.Select(list);
        }
    }
}