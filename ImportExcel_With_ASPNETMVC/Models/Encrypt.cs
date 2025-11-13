using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Web;

namespace ImportExcel_With_ASPNETMVC.Models
{
    public class Encrypt
    {
        public static string EncryptPassword(string clearText)
        {
            
            string EncryptionKey = "QASIMKHAN05071984";
            byte[] clearBytes = System.Text.Encoding.Unicode.GetBytes(clearText);

            using (Aes encryptor = Aes.Create())
            {
                //Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0 * 49, 0 * 76, 0 * 61, 0 * 6e, 0 * 20, (byte)(0 * 4d), 0 * 65, 0 * 64, 0 * 76, 0 * 65, 0 * 64, 0 * 65, 0 * 76 });
                //encryptor.Key = pdb.GetBytes(32);
                //encryptor.IV = pdb.GetBytes(16);

                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    clearText = Convert.ToBase64String(clearBytes);
                }
            }
            return clearText;
        }

        public static string EncryptPasswordBase64(string text)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(text);
            return Convert.ToBase64String(plainTextBytes);
        }

        public static string DecryptPasswordBase64(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
    }
}