using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace docx2tex.UI
{
    static class UserConfigHandler
    {
        private static Configuration _config;
        private static RecentConversionSection _recentConvs;

        public static void LoadConfiguration()
        {
            _config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);

            _recentConvs = _config.GetSection("recentConversions") as RecentConversionSection;
            if (_recentConvs == null)
            {
                _recentConvs = new RecentConversionSection();

                //trick:
                _recentConvs.SectionInformation.AllowExeDefinition = ConfigurationAllowExeDefinition.MachineToLocalUser;

                _config.Sections.Add("recentConversions", _recentConvs);
            }
        }

        public static List<FromToElement> GetRecentConversions()
        {
            List<FromToElement> ftes = new List<FromToElement>();
            foreach(FromToElement fte in _recentConvs.FromTos)
            {
                ftes.Add(fte);
            }

            return ftes;
        }

        public static void UpdateRecentConversion(List<FromToElement> ftes)
        {
            while (_recentConvs.FromTos.Count > 0)
            {
                _recentConvs.FromTos.RemoveAt(0);
            }

            for (int i = 0; i < ftes.Count && i < 10; i++)
            {
                FromToElement fte = ftes[i];
                fte.Order = i;
                _recentConvs.FromTos.Add(fte);
            }
        }

        public static void SaveConfiguration()
        {
            _config.Save();
        }
    }

    #region Configuration Data stores

	public class RecentConversionSection : ConfigurationSection
    {
        static RecentConversionSection()
        {
            _propFromTos = new ConfigurationProperty(
                "",
                typeof(FromToElementCollection),
                null,
                ConfigurationPropertyOptions.IsRequired | ConfigurationPropertyOptions.IsDefaultCollection
                );

            _properties = new ConfigurationPropertyCollection();

            _properties.Add(_propFromTos);
        }

        private static ConfigurationPropertyCollection _properties;
        private static ConfigurationProperty _propFromTos;

        public FromToElementCollection FromTos
        {
            get { return (FromToElementCollection)base[_propFromTos]; }
        }

        protected override ConfigurationPropertyCollection Properties
        {
            get
            {
                return _properties;
            }
        }
    }

    public class FromToElement : ConfigurationElement
    {
        static FromToElement()
        {
            _orderName = new ConfigurationProperty(
                "order",
                typeof(int),
                null,
                ConfigurationPropertyOptions.IsRequired
            );

            _fromName = new ConfigurationProperty(
                "from",
                typeof(string),
                null,
                ConfigurationPropertyOptions.IsRequired
            );

            _toName = new ConfigurationProperty(
                "to",
                typeof(string),
                null,
                ConfigurationPropertyOptions.IsRequired
            );

            _properties = new ConfigurationPropertyCollection();

            _properties.Add(_orderName);
            _properties.Add(_fromName);
            _properties.Add(_toName);
        }

        private static ConfigurationProperty _orderName;
        private static ConfigurationProperty _fromName;
        private static ConfigurationProperty _toName;

        private static ConfigurationPropertyCollection _properties;

        [ConfigurationProperty("order", IsRequired = true)]
        public int Order
        {
            get { return (int)base[_orderName]; }
            set { base[_orderName] = value; }
        }

        [ConfigurationProperty("from", IsRequired = true)]
        public string From
        {
            get { return (string)base[_fromName]; }
            set { base[_fromName] = value; }
        }

        [ConfigurationProperty("to", IsRequired = true)]
        public string To
        {
            get { return (string)base[_toName]; }
            set { base[_toName] = value; }
        }

        protected override ConfigurationPropertyCollection Properties
        {
            get { return _properties; }
        }
    }

    [ConfigurationCollection(typeof(FromToElement), AddItemName = "fromto", CollectionType = ConfigurationElementCollectionType.BasicMap)]
    public class FromToElementCollection : ConfigurationElementCollection
    {
        public override ConfigurationElementCollectionType CollectionType
        {
            get { return ConfigurationElementCollectionType.BasicMap; }
        }

        protected override string ElementName
        {
            get { return "fromto"; }
        }

        protected override ConfigurationPropertyCollection Properties
        {
            get
            {
                return new ConfigurationPropertyCollection();
            }
        }

        protected override ConfigurationElement CreateNewElement()
        {
            return new FromToElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return (element as FromToElement).Order.ToString();
        }

		public FromToElement this[int index]
		{
			get
			{
				return (FromToElement)base.BaseGet(index);
			}
			set
			{
				if (base.BaseGet(index) != null)
				{
					base.BaseRemoveAt(index);
				}
				base.BaseAdd(index, value);
			}
		}

		public new FromToElement this[string key]
		{
			get
			{
				return (FromToElement)base.BaseGet(key);
			}
		}

		public void Add(FromToElement item)
		{
			base.BaseAdd(item);
		}

        public void Remove(FromToElement item)
		{
			base.BaseRemove(item);
		}

		public void RemoveAt(int index)
		{
			base.BaseRemoveAt(index);
		}
    }

	#endregion
}
