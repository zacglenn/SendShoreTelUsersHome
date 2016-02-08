using System;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Configuration;
using System.Data.OleDb;

namespace SendShoreTelUsersHome
{
    public class AES
    {
        private static string key;

        public AES()
        {
            RetrieveKey();
        }

        static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV)
        {
            // Check arguments. 
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("Key");
            byte[] encrypted;
            // Create an Rijndael object 
            // with the specified key and IV. 
            using (Rijndael rijAlg = Rijndael.Create())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform. 
                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for encryption. 
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {

                            //Write all data to the stream. 
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            // Return the encrypted bytes from the memory stream. 
            return encrypted;
        }

        /// <summary>
        /// Decrypts the specified base64 encoded cipher text.
        /// </summary>
        /// <param name="base64EncodedCipherText">The base64 encoded cipher text.</param>
        /// <returns></returns>
        public static string Decrypt(string base64EncodedCipherText)
        {
            using (var alg = new RijndaelManaged())
            {
                alg.BlockSize = 256;
                alg.Key = System.Text.Encoding.ASCII.GetBytes(key).Take(32).ToArray();
                alg.Mode = CipherMode.ECB;
                alg.Padding = PaddingMode.Zeros;
                alg.IV = new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

                var cipherText = Convert.FromBase64String(base64EncodedCipherText);

                using (ICryptoTransform decryptor = alg.CreateDecryptor())
                {
                    using (var ms = new MemoryStream(cipherText))
                    {
                        using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                        {
                            using (var sr = new StreamReader(cs))
                            {
                                return sr.ReadToEnd().Replace("\0", string.Empty);
                            }
                        }
                    }
                }
            }
        }

        private void RetrieveKey()
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            AppSettingsSection globalSettingSection = (AppSettingsSection)config.GetSection("globalSettings");
            key = globalSettingSection.Settings["key"].Value;
        }

        public static string DecryptConnectionString(string encryptedConnectionString)
        {
            var connectionString = string.Empty;

            OleDbConnectionStringBuilder DBConfig = new OleDbConnectionStringBuilder(encryptedConnectionString);

            return connectionString =
                "Data Source=" + Decrypt(DBConfig["datasource"].ToString())
                + ";User ID=" + Decrypt(DBConfig["userid"].ToString())
                + ";Password=" + Decrypt(DBConfig["password"].ToString());
        }
    }
}
