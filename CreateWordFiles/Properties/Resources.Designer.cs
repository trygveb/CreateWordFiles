﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CreateWordFiles.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("CreateWordFiles.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to welcome;Welcome
        ///to;to
        ///Friday;Friday
        ///Saturday;Saturday
        ///Sunday;Sunday
        ///Monday;Monday
        ///weekend_member_fees;Members: SEK {0}/pass, all passes SEK {1}
        ///weekend_non_member_fees;Non members: SEK {0}/pass, all passes SEK {1}
        ///one_pass_sunday;(sunday counts as 2 passes)
        ///pg_pay;N/A
        ///weekend_swish_pay;N/A
        ///.
        /// </summary>
        internal static string texts_en {
            get {
                return ResourceManager.GetString("texts_en", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to welcome;Välkommen
        ///to;till
        ///Friday;Fredag
        ///Saturday;Lördag
        ///Sunday;Söndag
        ///Monday;Måndag
        ///weekend_member_fees;Medlem: {0} kr/pass, samtliga pass {1} kr
        ///weekend_non_member_fees;Ej medlem: {0} kr/pass, samtliga pass {1} kr
        ///one_pass_sunday;(söndag räknas som 2 pass)
        ///pg_pay;Betala gärna i förväg på PlusGiro 85 56 69-8 (MOTIV8&apos;S)
        ///weekend_swish_pay;Swisha till 070-422 82 27 (Arne G) eller kontanter ”i dörren” går också bra
        ///.
        /// </summary>
        internal static string texts_se {
            get {
                return ResourceManager.GetString("texts_se", resourceCulture);
            }
        }
    }
}
