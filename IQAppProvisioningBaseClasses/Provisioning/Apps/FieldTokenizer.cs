using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;
using IQAppProvisioningBaseClasses.Utility;

namespace IQAppProvisioningBaseClasses.Provisioning
{
    public class FieldTokenizer
    {
        public static string DoTokenSubstitutionsAndCleanSchema(ClientContext ctx, Field field)
        {
            return DoTokenSubstitutionsAndCleanSchema(ctx, ctx.Web, field);
        }
        public static string DoTokenSubstitutionsAndCleanSchema(ClientContext ctx, Web web, Field field)
        {
            var schemaXml = field.SchemaXml;
            var newXml = SubstituteGroupTokens(ctx, web, schemaXml);
            newXml = TokenizeLookupField(ctx, web, newXml);
            newXml = TokenizeTaxonomyField(ctx, field, newXml);
            newXml = newXml.RemoveXmlAttribute("Version");
            newXml = newXml.RemoveXmlAttribute("Sealed");
            newXml = newXml.RemoveXmlAttribute("SourceId");
            return newXml;
        }

        public static string DoTokenReplacement(ClientContext ctx, string schemaXml)
        {
            return DoTokenReplacement(ctx, ctx.Web, schemaXml);
        }

        public static string DoTokenReplacement(ClientContext ctx, Web web, string schemaXml)
        {
            var newXml = ReplaceGroupTokens(ctx, web, schemaXml);
            newXml = ReplaceListTokens(ctx, web, newXml);
            newXml = ReplaceTaxonomyTokens(ctx, web, newXml);
            return newXml;
        }
        
        private static string SubstituteGroupTokens(ClientContext ctx, Web web, string schemaXml)
        {
            if (!schemaXml.Contains("UserSelectionScope"))
            {
                return schemaXml;
            }
            var groupId = schemaXml.GetXmlAttribute("UserSelectionScope");
            if (groupId == null)
            {
                return schemaXml;
            }

            int id;
            if (int.TryParse(groupId, out id) && id != 0)
            {
                var group = web.SiteGroups.GetById(id);
                ctx.Load(group, g => g.Title);
                ctx.Load(web.AssociatedMemberGroup, g => g.Id);
                ctx.Load(web.AssociatedOwnerGroup, g => g.Id);
                ctx.Load(web.AssociatedVisitorGroup, g => g.Id);
                ctx.ExecuteQueryRetry();

                var tokenTitle = group.Title;

                if (id == web.AssociatedMemberGroup.Id)
                {
                    tokenTitle = "AssociatedMemberGroup";
                }
                if (id == web.AssociatedOwnerGroup.Id)
                {
                    tokenTitle = "AssociatedOwnerGroup";
                }
                if (id == web.AssociatedVisitorGroup.Id)
                {
                    tokenTitle = "AssociatedVisitorGroup";
                }

                schemaXml = schemaXml.SetXmlAttribute("UserSelectionScope", tokenTitle);
            }
            return schemaXml;
        }

        private static string TokenizeLookupField(ClientContext ctx, Web web, string schemaXml)
        {
            var retval = schemaXml;

            var fieldType = retval.GetXmlAttribute("Type");
            if (fieldType == "Lookup")
            {
                List lookupTarget = null;
                var listIdOrUrl = retval.GetXmlAttribute("List");
                if (listIdOrUrl != null)
                {
                    Guid listGuid;
                    if (Guid.TryParse(listIdOrUrl, out listGuid))
                    {
                        lookupTarget = web.Lists.GetById(listGuid);
                        ctx.Load(lookupTarget, l => l.Title);
                        ctx.ExecuteQueryRetry();
                    }
                    else if (listIdOrUrl.Contains("/"))
                    {
                        if (!listIdOrUrl.StartsWith("/"))
                        {
                            listIdOrUrl = "/" + listIdOrUrl;
                        }
                        var baseUrl = web.ServerRelativeUrl == "/" ? "" : web.ServerRelativeUrl;

                        //Get list is new since CSOM v15.0.4701.1001
                        if (ctx.ServerVersion >= Version.Parse("15.0.4701.1001"))
                        {
                            lookupTarget = web.GetList(baseUrl + listIdOrUrl);
                            ctx.Load(lookupTarget, l => l.Title);
                            ctx.ExecuteQueryRetry();
                        }
                        else
                        {
                            var lists = web.Lists;
                            ctx.Load(lists, ls => ls.Include(l => l.DefaultViewUrl, l => l.Title));
                            ctx.ExecuteQueryRetry();
                            foreach (var l in lists)
                            {
                                if (l.DefaultViewUrl.ToLower().Contains(listIdOrUrl.ToLower()))
                                {
                                    lookupTarget = l;
                                    break;
                                }
                            }
                        }
                    }
                    if (lookupTarget != null && lookupTarget.IsPropertyAvailable("Title"))
                    {
                        retval = retval.SetXmlAttribute("List", "{@ListId:" + lookupTarget.Title + "}");
                    }
                }
            }

            return retval;
        }

        private static string TokenizeTaxonomyField(ClientContext ctx, Field field, string schemaXml)
        {
            var fieldType = schemaXml.GetXmlAttribute("Type");
            if (!fieldType.StartsWith("TaxonomyField")) return schemaXml;

            schemaXml = schemaXml.RemoveXmlAttribute("List");
            schemaXml = schemaXml.RemoveXmlAttribute("WebId");
            schemaXml = schemaXml.RemoveXmlAttribute("SourceID");
            schemaXml = schemaXml.RemoveXmlAttribute("Version");
            //Default Value

            var taxonomyField = ctx.CastTo<TaxonomyField>(field);
            ctx.Load(taxonomyField);
            ctx.ExecuteQuery();

            var defaultValue = taxonomyField.DefaultValue;
            if (!string.IsNullOrEmpty(defaultValue))
            {
                schemaXml = schemaXml.Replace(defaultValue, "");
            }

            var sspId = taxonomyField.SspId.ToString();
            schemaXml = schemaXml.Replace(sspId, "{@SspId}");

            if (taxonomyField.TermSetId != default(Guid) ||
                taxonomyField.AnchorId != default(Guid))
            {
                var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                var termStore = TermStoreUtility.GetTermStore(ctx, taxonomySession);
                TermSet termSet = null;
                Term anchorTerm = null;

                if (taxonomyField.TermSetId != default(Guid))
                {
                    termSet = termStore.GetTermSet(taxonomyField.TermSetId);
                    ctx.Load(termSet, ts => ts.Name);
                }
                if (taxonomyField.AnchorId != default(Guid))
                {
                    anchorTerm = termStore.GetTerm(taxonomyField.AnchorId);
                    ctx.Load(anchorTerm, t => t.Name);
                }
                try
                {
                    ctx.ExecuteQuery();
                }
                catch
                {
                    //ignore
                }

                if (termSet != null && termSet.IsPropertyAvailable("Name"))
                {
                    schemaXml = schemaXml.Replace(taxonomyField.TermSetId.ToString(), $"{{@TermSet:{termSet.Name}}}");
                }
                else
                {
                    schemaXml = schemaXml.Replace(taxonomyField.TermSetId.ToString(), $"00000000-0000-0000-0000-000000000000");
                }
                if (anchorTerm != null && anchorTerm.IsPropertyAvailable("Name"))
                {
                    schemaXml = schemaXml.Replace(taxonomyField.AnchorId.ToString(), $"{{@AnchorTermId:{anchorTerm.Name}}}");
                }
                else
                {
                    schemaXml = schemaXml.Replace(taxonomyField.AnchorId.ToString(), $"00000000-0000-0000-0000-000000000000");
                }
            }

            return schemaXml;
        }
        private static string ReplaceGroupTokens(ClientContext ctx, Web web, string schemaXml)
        {
            if (!schemaXml.Contains("UserSelectionScope"))
            {
                return schemaXml;
            }
            if (web.AppInstanceId != default(Guid))
            {
                return schemaXml.RemoveXmlAttribute("UserSelectionScope");
            }
            var groupName = schemaXml.GetXmlAttribute("UserSelectionScope");
            if (groupName == null || groupName == "0")
            {
                return schemaXml;
            }

            Group group;
            if (groupName == "AssociatedMemberGroup")
            {
                group = web.AssociatedMemberGroup;
            }
            else if (groupName == "AssociatedOwnerGroup")
            {
                group = web.AssociatedOwnerGroup;
            }
            else if (groupName == "AssociatedVisitorGroup")
            {
                group = web.AssociatedVisitorGroup;
            }
            else
            {
                group = web.SiteGroups.GetByName(groupName);
            }
            ctx.Load(group, g => g.Id);
            ctx.ExecuteQueryRetry();

            schemaXml = schemaXml.SetXmlAttribute("UserSelectionScope", group.Id.ToString());

            return schemaXml;
        }

        private static string ReplaceListTokens(ClientContext ctx, Web web, string schemaXml)
        {
            var retval = schemaXml;

            var listTitle = retval.GetInnerText("{@ListId:", "}", true);
            if (!string.IsNullOrEmpty(listTitle))
            {
                var list = web.Lists.GetByTitle(listTitle);
                ctx.Load(list, l => l.Id);

                try
                {
                    ctx.ExecuteQueryRetry();
                }
                catch
                {
                    //Ignore. In some versions of CSOM the list not existing will give a runtime error
                    //In others it just doesn't load the object and we can check property availabilty
                }
                if (list.IsPropertyAvailable("Id"))
                {
                    retval = retval.SetXmlAttribute("List", "{" + list.Id + "}");
                }
            }

            return retval;
        }

        private static string ReplaceTaxonomyTokens(ClientContext ctx, Web web, string schemaXml)
        {
            var tokens = new List<string>()
            {
                "{@DefaultValue:",
                "{@SspId}",
                "{@TermSet:",
                "{@AnchorTermId:"
            };

            var foundTokens = tokens.Where(t => schemaXml.Contains(t)).ToList();
            if (foundTokens.Count == 0) return schemaXml;

            var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            var termStore = TermStoreUtility.GetTermStore(ctx, taxonomySession);
            ctx.Load(termStore, ts => ts.Id);
            ctx.ExecuteQuery();

            schemaXml = schemaXml.Replace("{@SspId}", termStore.Id.ToString());
            var termSetName = schemaXml.GetInnerText("{@TermSet:", "}");
            if (!string.IsNullOrEmpty(termSetName))
            {
                var termSets = termStore.GetTermSetsByName(termSetName, (int)web.Language);
                ctx.Load(termSets, ts => ts.Include(t => t.Id, t => t.Name));
                ctx.ExecuteQuery();

                if (termSets.Count == 0)
                {
                    throw new InvalidOperationException($"Unable to find term set {termSetName}.");
                }

                var termSet = GetCorrectTermSet(ctx, schemaXml, termSets);

                System.Diagnostics.Debug.WriteLine($"{schemaXml.GetXmlAttribute("DisplayName")} | {termSet.Id} | {termSet.Name}");

                schemaXml = schemaXml.Replace($"{{@TermSet:{termSetName}}}", termSet.Id.ToString());
                var terms = termSet.GetAllTerms();
                ctx.Load(terms);
                ctx.ExecuteQuery();
                if (foundTokens.Contains("{@AnchorTermId:"))
                {
                    var anchorTermName = schemaXml.GetInnerText("{@AnchorTermId:", "}");
                    var foundAnchorTerm = terms.FirstOrDefault(t => t.Name == anchorTermName);
                    schemaXml = schemaXml.Replace($"{{@AnchorTermId:{anchorTermName}}}", foundAnchorTerm?.Id.ToString() ?? "");
                }

                if (foundTokens.Contains("{@DefaultValue:"))
                {
                    var defaultValueTermName = schemaXml.GetInnerText("{@DefaultValue:", "}");
                    var foundDefaultTerm = terms.FirstOrDefault(t => t.Name == defaultValueTermName);
                    schemaXml = schemaXml.Replace($"{{@DefaultValue:{defaultValueTermName}}}", foundDefaultTerm != null ? $"-1;#{defaultValueTermName}|{foundDefaultTerm.Id.ToString()}" : "");
                }
            }

            return schemaXml;
        }

        private static TermSet GetCorrectTermSet(ClientContext ctx, string schemaXml, TermSetCollection termSets)
        {
            //This is a challenge when there is more than one 
            if (termSets.Count == 1) return termSets[0];
            if (termSets.Count == 0) return null;

            var anchorTermName = schemaXml.GetInnerText("{@AnchorTermId:", "}");
            var defaultValueTermName = schemaXml.GetInnerText("{@DefaultValue:", "}");
            TermSet bestMatch = null;
            var maxTermCount = 0;

            foreach (var termSet in termSets)
            {
                var terms = termSet.GetAllTerms();
                ctx.Load(terms);
                ctx.Load(termSet.Group, g => g.Name);
                try
                {
                    ctx.ExecuteQuery();

                    if (!termSet.Group.Name.StartsWith("Site Collection"))
                    {

                        if (!string.IsNullOrEmpty(anchorTermName) && !string.IsNullOrEmpty(defaultValueTermName))
                        {
                            if (terms.FirstOrDefault(t => t.Name == anchorTermName) != null && terms.FirstOrDefault(t => t.Name == defaultValueTermName) != null)
                            {
                                return termSet;
                            }
                        }
                        else if (!string.IsNullOrEmpty(anchorTermName))
                        {
                            if (terms.FirstOrDefault(t => t.Name == anchorTermName) != null)
                            {
                                return termSet;
                            }
                        }
                        else if (terms.FirstOrDefault(t => t.Name == defaultValueTermName) != null)
                        {
                            return termSet;
                        }
                        else
                        {
                            if (terms.Count > maxTermCount || bestMatch == null)
                            {
                                maxTermCount = terms.Count;
                                bestMatch = termSet;
                            }
                        }
                    }
                }
                catch
                {
                    //ignored
                }
            }
            return bestMatch;
        }
    }
}