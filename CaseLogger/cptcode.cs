using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CaseLogger
{
    class cptcode
    {
        public string code{ get; set; }
        public string description{ get; set; }
        public string area { get; set; }
        public string type { get; set; }
        public List<string> descriptionwords;

        public cptcode()
        {
        }

        public List<string> words
        {
            get { return descriptionwords; }
            set { descriptionwords = value; }
        }

        public void generatewords()
        {
            char[] delimiterch = { ' ' };
            description = StringWordsRemove(description.ToLower()); 
            string[] words = description.Split(delimiterch);
            descriptionwords = new List<string>(words);
        }

        public static string StringWordsRemove(string stringToClean)
        {
            List<string> wordstoremove = "and or but of with for the an in using eg".Split(' ').ToList();
            return string.Join(" ", stringToClean.Split(new[] { ' ', ',', '.', '?', '!',';',':','(',')' }, StringSplitOptions.RemoveEmptyEntries).Except(wordstoremove));
        }
        
    }
}
