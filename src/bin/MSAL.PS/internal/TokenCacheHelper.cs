using System;
using System.IO;
using System.Security.Cryptography;
using Microsoft.Identity.Client;

public static class TokenCacheHelper
{
    public static void EnableSerialization(ITokenCache tokenCache)
    {
        tokenCache.SetBeforeAccess(BeforeAccessNotification);
        tokenCache.SetAfterAccess(AfterAccessNotification);
    }

    /// <summary>
    /// Path to the token cache
    /// </summary>
    public static readonly string CacheFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MSAL.PS", "MSAL.PS.msalcache.bin3");

    private static readonly object FileLock = new object();

    private static void BeforeAccessNotification(TokenCacheNotificationArgs args)
    {
        lock (FileLock)
        {
            args.TokenCache.DeserializeMsalV3(File.Exists(CacheFilePath)
                    ? ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath),
                                              null,
                                              DataProtectionScope.CurrentUser)
                    : null);
        }
    }

    private static void AfterAccessNotification(TokenCacheNotificationArgs args)
    {
        // if the access operation resulted in a cache update
        if (args.HasStateChanged)
        {
            lock (FileLock)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(CacheFilePath));
                // reflect changes in the persistent store
                File.WriteAllBytes(CacheFilePath,
                                    ProtectedData.Protect(args.TokenCache.SerializeMsalV3(),
                                                            null,
                                                            DataProtectionScope.CurrentUser)
                                    );
            }
        }
    }
}
