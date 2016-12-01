namespace IQAppProvisioningBaseClasses.Provisioning
{
    public class ComposedLook
    {
        /// <summary>
        /// Theme Title and Name
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Server relative URL of the .spcolor file
        /// Should be in the root web of the site collection
        /// </summary>
        public string ThemeUrl { get; set; }

        /// <summary>
        /// Server relative URL of the URL of the FontFile
        /// Should be in the root web of the site collection
        /// </summary>
        public string FontSchemeURL { get; set; }

        /// <summary>
        /// Server relative URL of the Background Image
        /// Should be in the root web of the site collection 
        /// </summary>
        public string ImageUrl { get; set; }
    }
}
