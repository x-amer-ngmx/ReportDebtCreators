using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using ReportDebtCreators.model;

namespace ReportDebtCreators.enginer
{
    public static class Global
    {
        public static IEnumerable<FileInfo> GetFilesByExtensions(this DirectoryInfo dir, params string[] extensions)
        {
            if (extensions == null)
                throw new ArgumentNullException("extensions");
            IEnumerable<FileInfo> files = dir.EnumerateFiles();
            return files.Where(f => extensions.Contains(f.Extension));
        }

        public static IList<StructExelModel> GetRangePack(this IList<StructExelModel> fromPack,DateTime? f, DateTime? t)
        {
            return (from p in fromPack
             orderby p.DateIndex descending
             where (p.DateIndex > f && p.DateIndex <= t)
             select p).ToList();
        }
    }
}
