// ============================================================================
// 18 - Master model - create or update perspective models.cs
//
// Based on the idea of m-kovalsky/Tabular MasterModel.cs, rebuilt for TE3 and
// PBIP: one master semantic model with perspectives -> one smaller semantic
// model per perspective (e.g. for conversational BI / Copilot / data agents).
// No report is created or copied.
//
// Run it with the MASTER model open in Tabular Editor 3. The master itself is
// never changed.
//
// Workflow:
//   1. Pick the output folder (the folder that holds the master
//      .SemanticModel folder). The last used folder is remembered.
//   2. A grid lists every perspective. Tick the ones to process and edit the
//      model name (default "<Perspective> <Master name>"). The action follows
//      the name:
//        Create  - no <Name>.SemanticModel folder exists yet
//        Update  - the folder exists (models created by this macro are found
//                  automatically, whatever their name)
//   3. Create writes <Name>.SemanticModel\ with definition\ (TMDL),
//      definition.pbism and .platform next to the master.
//   4. Update merges master changes into the existing model WITHOUT
//      overwriting local work (three-way compare, see below).
//
// Building a perspective model:
//   - Keeps the tables, columns, measures and hierarchies in the perspective.
//   - Also keeps, HIDDEN, everything they need: DAX dependencies (incl. DAX
//     UDFs), sort-by columns, hierarchy levels, relationships between kept
//     tables and their key columns, columns used in RLS filters, and all
//     columns of calculated tables / calculation groups.
//   - Removes all perspectives, unused DAX UDFs, translations and role
//     permissions for removed objects.
//   - Applies the original MasterModel.cs annotations for the perspective
//     ({P} = perspective name) and removes them from the output:
//       {P}_Expression           measure: replace expression
//       {P}_Hide / {P}_Unhide     table/column/measure: IsHidden true/false
//       {P}_Remove                data source / role / partition: delete
//       {P}_RemoveKPI             measure: delete KPI
//       {P}_RLS_<Table>           role: RLS filter expression for <Table>
//       {P}_UpdateRoleMembers     role: members, separated by , or ;
//       {P}_ModelPermission       role: Administrator or Read
//       {P}_UpdateDataSource      table: legacy data source for its partitions
//       {P}_UpdatePartitionQuery  table: "old|new" text replacement in the
//                                 partition query/M, plus %Key% placeholders
//                                 replaced by the partition's annotations
//
// Updating (three-way compare):
//   Every generated model stores a fingerprint of each object as it was
//   generated (model annotation MasterModel_Baseline). On update the
//   perspective model is generated again in memory and compared per table,
//   column, measure, hierarchy, partition, calculation item, relationship,
//   role and shared expression:
//     New in master                      -> Add
//     Changed in master only             -> Take master
//     Changed locally only               -> Keep mine
//     Changed in both / no fingerprint   -> YOU CHOOSE (Take master / Keep mine)
//     Removed from master                -> YOU CHOOSE (Remove / Keep mine)
//     Deleted locally                    -> Keep deleted (or Add)
//     Objects that only exist locally    -> never touched
//   A compare panel shows the master and local definition side by side.
//   Models without a fingerprint (created before this macro) can be updated
//   too: every difference is then asked.
//   The previous definition folder is backed up to %TEMP%\MasterModelBackup.
//
// IMPORTANT:
//   - The generated/updated models are written to disk; they are not opened.
//     Open <Name>.SemanticModel\definition in Tabular Editor, or deploy /
//     sync the folder with Fabric Git integration. Power BI Desktop needs a
//     report (.pbip) to open a semantic model.
//   - Linguistic schema / Copilot "Prep data for AI" settings are copied as-is
//     and may mention removed objects - review them for the smaller model.
// ============================================================================

#r "System.Drawing"

using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using TOM = Microsoft.AnalysisServices.Tabular;

ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

const string AnnSource = "MasterModel_Source";
const string AnnPerspective = "MasterModel_Perspective";
const string AnnUpdated = "MasterModel_UpdatedAt";
const string AnnBaseline = "MasterModel_Baseline";

class DerivedModel
{
    public string Folder;          // ...\Name.SemanticModel
    public string Name;
    public string Perspective;     // from MasterModel_Perspective (null if not created by this macro)
    public string Source;
    public bool HasPerspectives;   // looks like a master model - never offered as update target
}

class PerspectiveRow
{
    public string Perspective;
    public int Tables, Columns, Measures;
}

class Unit
{
    public string Key;
    public string Kind;
    public string Table;
    public string Name;
    public TOM.MetadataObject Obj;
    public string Hash;
    public string Label { get { return Kind == "Table" || Kind == "Role" || Kind == "Expression" || Kind == "Relationship" ? Name : Table + "[" + Name + "]"; } }
}

class MergeItem
{
    public string Key;
    public string Kind;
    public string Label;
    public string Change;
    public string[] Options;
    public string Action;          // null = needs a decision
    public Unit Master;
    public Unit Local;
}

// ---------------------------------------------------------------------------
// General helpers
// ---------------------------------------------------------------------------
bool SameName(string a, string b) { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }

// Every TOM type has its own annotation collection (MeasureAnnotationCollection,
// TableAnnotationCollection, ...), so they are accessed through reflection.
object Annotations(object obj)
{
    var prop = obj == null ? null : obj.GetType().GetProperty("Annotations");
    return prop == null ? null : prop.GetValue(obj, null);
}
List<TOM.Annotation> AnnotationList(object obj)
{
    var coll = Annotations(obj) as System.Collections.IEnumerable;
    return coll == null ? new List<TOM.Annotation>() : coll.Cast<TOM.Annotation>().ToList();
}
string GetAnn(object obj, string name)
{
    var coll = Annotations(obj);
    if (coll == null) return null;
    var ann = coll.GetType().GetMethod("Find", new[] { typeof(string) }).Invoke(coll, new object[] { name }) as TOM.Annotation;
    return ann == null ? null : ann.Value;
}
void SetAnn(object obj, string name, string value)
{
    var coll = Annotations(obj);
    if (coll == null) return;
    var ann = coll.GetType().GetMethod("Find", new[] { typeof(string) }).Invoke(coll, new object[] { name }) as TOM.Annotation;
    if (ann != null) ann.Value = value;
    else coll.GetType().GetMethod("Add", new[] { typeof(TOM.Annotation) }).Invoke(coll, new object[] { new TOM.Annotation { Name = name, Value = value } });
}
void RemoveAnn(object obj, TOM.Annotation ann)
{
    var coll = Annotations(obj);
    if (coll != null) coll.GetType().GetMethod("Remove", new[] { typeof(TOM.Annotation) }).Invoke(coll, new object[] { ann });
}

object CloneObject(object obj)
{
    // TOM subclasses re-declare Clone(); prefer the most derived one (avoids AmbiguousMatchException)
    var method = obj.GetType().GetMethods()
        .Where(m => m.Name == "Clone" && m.GetParameters().Length == 0)
        .OrderBy(m => m.DeclaringType == obj.GetType() ? 0 : 1)
        .First();
    return method.Invoke(obj, null);
}

void CopyDirectory(string source, string target)
{
    Directory.CreateDirectory(target);
    foreach (var file in Directory.GetFiles(source)) File.Copy(file, Path.Combine(target, Path.GetFileName(file)), true);
    foreach (var dir in Directory.GetDirectories(source)) CopyDirectory(dir, Path.Combine(target, Path.GetFileName(dir)));
}

// Calls XxxSerializer.SerializeObject(obj, <defaults>) regardless of the exact overload
string SerializeObject(Type serializer, object obj)
{
    var method = serializer.GetMethods()
        .Where(m => m.Name == "SerializeObject" && m.GetParameters().Length > 0
            && m.GetParameters()[0].ParameterType.IsAssignableFrom(obj.GetType()))
        .OrderBy(m => m.GetParameters().Length).First();
    var args = method.GetParameters().Select((p, i) => i == 0 ? obj
        : p.HasDefaultValue ? p.DefaultValue
        : p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null).ToArray();
    return (string)method.Invoke(null, args);
}

// Removes volatile/bookkeeping keys so equal definitions hash equally
var volatileKeys = new[] { "modifiedTime", "structureModifiedTime", "refreshedTime", "state", "errorMessage", "changedProperties", "lastProcessed", "lastUpdate", "lastSchemaUpdate" };
void Normalize(JToken token)
{
    var obj = token as JObject;
    if (obj != null)
    {
        foreach (var key in volatileKeys) obj.Remove(key);
        var annotations = obj["annotations"] as JArray;
        if (annotations != null)
        {
            foreach (var a in annotations.OfType<JObject>().Where(a => ((string)a["name"] ?? "").StartsWith("MasterModel_")).ToList()) a.Remove();
            if (annotations.Count == 0) obj.Remove("annotations");
        }
        foreach (var p in obj.Properties().ToList()) Normalize(p.Value);
    }
    var array = token as JArray;
    if (array != null) foreach (var item in array) Normalize(item);
}

string HashOf(TOM.MetadataObject obj, Action<JObject> tweak)
{
    var json = JObject.Parse(SerializeObject(typeof(TOM.JsonSerializer), obj));
    Normalize(json);
    if (tweak != null) tweak(json);
    using (var sha = System.Security.Cryptography.SHA1.Create())
    {
        var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(json.ToString(Newtonsoft.Json.Formatting.None)));
        return BitConverter.ToString(bytes).Replace("-", "").Substring(0, 16);
    }
}

string Readable(TOM.MetadataObject obj)
{
    if (obj == null) return "(does not exist)";
    try { return SerializeObject(typeof(TOM.TmdlSerializer), obj).Replace("\t", "    ").Replace("\r\n", "\n").Replace("\n", "\r\n"); }
    catch { return JObject.Parse(SerializeObject(typeof(TOM.JsonSerializer), obj)).ToString(); }
}

System.Collections.IEnumerable Functions(TOM.Model mdl)
{
    var prop = mdl.GetType().GetProperty("Functions");
    return prop == null ? null : prop.GetValue(mdl, null) as System.Collections.IEnumerable;
}

IEnumerable<object> AllObjects(TOM.Model mdl)
{
    var list = new List<object> { mdl };
    foreach (var t in mdl.Tables)
    {
        list.Add(t);
        list.AddRange(t.Columns); list.AddRange(t.Measures); list.AddRange(t.Hierarchies); list.AddRange(t.Partitions);
        if (t.CalculationGroup != null) list.AddRange(t.CalculationGroup.CalculationItems);
    }
    list.AddRange(mdl.Relationships); list.AddRange(mdl.Roles); list.AddRange(mdl.Expressions);
    list.AddRange(mdl.DataSources); list.AddRange(mdl.Perspectives); list.AddRange(mdl.Cultures);
    var functions = Functions(mdl);
    if (functions != null) list.AddRange(functions.Cast<object>());
    return list;
}

bool IsAttached(TOM.MetadataObject obj, TOM.Model mdl)
{
    var current = obj;
    while (current != null)
    {
        if (current == mdl) return true;
        current = current.Parent;
    }
    return false;
}

string RelationshipKey(TOM.SingleColumnRelationship r)
{
    return r.FromTable.Name + "[" + r.FromColumn.Name + "] -> " + r.ToTable.Name + "[" + r.ToColumn.Name + "]";
}

// ---------------------------------------------------------------------------
// Which master objects belong to a perspective (computed on the TE model,
// which knows DAX dependencies)
// ---------------------------------------------------------------------------
string KeyOf(object o)
{
    if (o is Measure) { var m = (Measure)o; return "Measure|" + m.Table.Name + "|" + m.Name; }
    if (o is Column) { var c = (Column)o; return "Column|" + c.Table.Name + "|" + c.Name; }
    if (o is Hierarchy) { var h = (Hierarchy)o; return "Hierarchy|" + h.Table.Name + "|" + h.Name; }
    if (o is Table) return "Table|" + ((Table)o).Name;
    var nameProp = o.GetType().GetProperty("Name");
    return o.GetType().Name + "|" + (nameProp == null ? o.GetHashCode().ToString() : (string)nameProp.GetValue(o, null));
}

// DependsOn in the TE object model: enumerates keys (dictionary) or objects
List<object> DependenciesOf(object o)
{
    var result = new List<object>();
    var prop = o.GetType().GetProperty("DependsOn");
    var deps = prop == null ? null : prop.GetValue(o, null);
    if (deps == null) return result;
    foreach (var name in new[] { "Measures", "Columns", "Tables" })
    {
        var sub = deps.GetType().GetProperty(name);
        var items = sub == null ? null : sub.GetValue(deps, null) as System.Collections.IEnumerable;
        if (items != null) result.AddRange(items.Cast<object>());
    }
    var all = deps as System.Collections.IEnumerable;
    if (all != null)
        foreach (var item in all)
        {
            var keyProp = item.GetType().GetProperty("Key");
            result.Add(keyProp != null ? keyProp.GetValue(item, null) : item);
        }
    return result.Where(x => x != null).Distinct().ToList();
}

void ComputePerspective(string perspective, HashSet<string> keep, HashSet<string> visible)
{
    var queue = new Queue<object>();
    Action<object, bool> enqueue = (o, isVisible) =>
    {
        if (o == null) return;
        var key = KeyOf(o);
        if (isVisible) visible.Add(key);
        if (keep.Add(key)) queue.Enqueue(o);
    };

    foreach (var t in Model.Tables)
    {
        if (!t.InPerspective[perspective]) continue;
        enqueue(t, true);
        var members = t.Columns.Where(c => c.InPerspective[perspective]).Cast<object>()
            .Concat(t.Measures.Where(m => m.InPerspective[perspective]))
            .Concat(t.Hierarchies.Where(h => h.InPerspective[perspective])).ToList();
        if (members.Count == 0) members = t.Columns.Cast<object>().Concat(t.Measures).Concat(t.Hierarchies).ToList();
        foreach (var member in members) enqueue(member, true);
    }

    bool changed = true;
    while (changed)
    {
        while (queue.Count > 0)
        {
            var o = queue.Dequeue();
            if (o is Measure) enqueue(((Measure)o).Table, false);
            if (o is Hierarchy)
            {
                var h = (Hierarchy)o;
                enqueue(h.Table, false);
                foreach (var level in h.Levels) enqueue(level.Column, false);
            }
            if (o is Column)
            {
                var c = (Column)o;
                enqueue(c.Table, false);
                enqueue(c.SortByColumn, false);
                foreach (var g in c.GroupByColumns) enqueue(g, false);
            }
            if (o is Table)
            {
                var t = (Table)o;
                var calendars = t.GetType().GetProperty("Calendars");
                var calendarList = calendars == null ? null : calendars.GetValue(t, null) as System.Collections.IEnumerable;
                if (t is CalculatedTable || t is CalculationGroupTable || (calendarList != null && calendarList.Cast<object>().Any()))
                    foreach (var c in t.Columns) enqueue(c, false);
                if (t is CalculationGroupTable)
                    foreach (var item in ((CalculationGroupTable)t).CalculationItems)
                        foreach (var dep in DependenciesOf(item)) enqueue(dep, false);
            }
            foreach (var dep in DependenciesOf(o)) enqueue(dep, false);   // DAX (measures, calc columns/tables, UDFs)
        }

        // Relationships between kept tables, and columns used in RLS filters
        changed = false;
        foreach (var r in Model.Relationships.OfType<SingleColumnRelationship>())
        {
            if (!keep.Contains("Table|" + r.FromTable.Name) || !keep.Contains("Table|" + r.ToTable.Name)) continue;
            if (keep.Add("Relationship|" + r.FromTable.Name + "[" + r.FromColumn.Name + "] -> " + r.ToTable.Name + "[" + r.ToColumn.Name + "]")) changed = true;
            enqueue(r.FromColumn, false);
            enqueue(r.ToColumn, false);
        }
        foreach (var role in Model.Roles)
            foreach (var t in Model.Tables.Where(t => keep.Contains("Table|" + t.Name)))
            {
                var filter = role.RowLevelSecurity[t];
                if (string.IsNullOrWhiteSpace(filter)) continue;
                foreach (var c in t.Columns)
                    if (filter.IndexOf("[" + c.Name + "]", StringComparison.OrdinalIgnoreCase) >= 0) enqueue(c, false);
            }
        if (queue.Count > 0) changed = true;
    }
}

// ---------------------------------------------------------------------------
// Build the perspective model as a separate TOM database (master untouched)
// ---------------------------------------------------------------------------
// The underlying TOM database of the open model. TE exposes it as Database.TOMDatabase;
// the model's MetadataObject is not public in TE3, so reflection is the fallback.
TOM.Database masterDb = null;
var tomDatabaseProp = Model.Database.GetType().GetProperty("TOMDatabase");
if (tomDatabaseProp != null) masterDb = tomDatabaseProp.GetValue(Model.Database, null) as TOM.Database;
if (masterDb == null)
{
    var metadataProp = Model.GetType().GetProperty("MetadataObject",
        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
    var tomModel = metadataProp == null ? null : metadataProp.GetValue(Model, null) as TOM.Model;
    if (tomModel != null) masterDb = tomModel.Database as TOM.Database;
}
if (masterDb == null) { Error("Could not access the TOM database of the open model."); return; }

var masterName = !string.IsNullOrWhiteSpace(masterDb.Name)
    && !Regex.IsMatch(masterDb.Name, "^[0-9a-fA-F-]{32,36}$") ? masterDb.Name : "Model";
var perspectiveNames = Model.Perspectives.Select(p => p.Name).ToList();

TOM.Database RoundTrip(TOM.Database db)
{
    // TMDL round trip so generated and on-disk models are compared in the same shape
    var temp = Path.Combine(Path.GetTempPath(), "MasterModel_" + Guid.NewGuid().ToString("N"));
    try
    {
        TOM.TmdlSerializer.SerializeDatabaseToFolder(db, temp);
        return TOM.TmdlSerializer.DeserializeDatabaseFromFolder(temp);
    }
    finally { try { Directory.Delete(temp, true); } catch { } }
}

TOM.Database BuildPerspectiveModel(string perspective, string newName, List<string> log)
{
    var keep = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var visible = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    ComputePerspective(perspective, keep, visible);

    // Independent copy of the master via TMDL (a JSON copy fails when the database name is empty, as with PBIP)
    var db = RoundTrip(masterDb);
    db.Name = newName;
    db.ID = Guid.NewGuid().ToString();
    var mdl = db.Model;
    var prefix = perspective + "_";

    // Relationships first (they reference columns), then roles' permissions, then objects
    foreach (var r in mdl.Relationships.OfType<TOM.SingleColumnRelationship>().ToList())
        if (!keep.Contains("Relationship|" + RelationshipKey(r))) mdl.Relationships.Remove(r);
    foreach (var r in mdl.Relationships.Where(r => !(r is TOM.SingleColumnRelationship)).ToList())
        if (!keep.Contains("Table|" + r.FromTable.Name) || !keep.Contains("Table|" + r.ToTable.Name)) mdl.Relationships.Remove(r);

    foreach (var t in mdl.Tables.ToList())
    {
        if (!keep.Contains("Table|" + t.Name))
        {
            foreach (var role in mdl.Roles)
                foreach (var tp in role.TablePermissions.Where(tp => tp.Table == t).ToList()) role.TablePermissions.Remove(tp);
            mdl.Tables.Remove(t);
            continue;
        }
        foreach (var h in t.Hierarchies.ToList()) if (!keep.Contains("Hierarchy|" + t.Name + "|" + h.Name)) t.Hierarchies.Remove(h);
        foreach (var m in t.Measures.ToList()) if (!keep.Contains("Measure|" + t.Name + "|" + m.Name)) t.Measures.Remove(m);
        foreach (var c in t.Columns.Where(c => c.Type != TOM.ColumnType.RowNumber).ToList())
        {
            if (keep.Contains("Column|" + t.Name + "|" + c.Name)) continue;
            foreach (var other in t.Columns.Where(o => o.SortByColumn == c).ToList()) other.SortByColumn = null;
            foreach (var role in mdl.Roles)
                foreach (var tp in role.TablePermissions)
                    foreach (var cp in tp.ColumnPermissions.Where(cp => cp.Column == c).ToList()) tp.ColumnPermissions.Remove(cp);
            t.Columns.Remove(c);
        }
        // A table kept only for a table-level DAX reference still needs a column
        if (!t.Columns.Any(c => c.Type != TOM.ColumnType.RowNumber))
            log.Add("Table '" + t.Name + "' has no kept columns - check the perspective.");
    }

    // Variations pointing to removed relationships/tables (auto date/time)
    foreach (var t in mdl.Tables)
        foreach (var c in t.Columns)
            foreach (var v in c.Variations.ToList())
                if (v.Relationship == null || !IsAttached(v.Relationship, mdl)) c.Variations.Remove(v);

    // Unused DAX UDFs
    var functions = Functions(mdl);
    if (functions != null)
    {
        var removeMethod = functions.GetType().GetMethods().FirstOrDefault(m => m.Name == "Remove" && m.GetParameters().Length == 1 && !m.GetParameters()[0].ParameterType.Equals(typeof(string)));
        foreach (var f in functions.Cast<object>().ToList())
        {
            var fname = (string)f.GetType().GetProperty("Name").GetValue(f, null);
            if (!keep.Contains("Function|" + fname) && removeMethod != null) removeMethod.Invoke(functions, new[] { f });
        }
    }

    foreach (var p in mdl.Perspectives.ToList()) mdl.Perspectives.Remove(p);

    // Objects kept only because something depends on them are hidden
    foreach (var t in mdl.Tables)
    {
        if (!visible.Contains("Table|" + t.Name)) t.IsHidden = true;
        foreach (var c in t.Columns.Where(c => c.Type != TOM.ColumnType.RowNumber))
            if (!visible.Contains("Column|" + t.Name + "|" + c.Name)) c.IsHidden = true;
        foreach (var m in t.Measures) if (!visible.Contains("Measure|" + t.Name + "|" + m.Name)) m.IsHidden = true;
        foreach (var h in t.Hierarchies) if (!visible.Contains("Hierarchy|" + t.Name + "|" + h.Name)) h.IsHidden = true;
    }

    // --- Original MasterModel.cs annotations for this perspective ---
    foreach (var t in mdl.Tables)
    {
        if (GetAnn(t, prefix + "Unhide") != null) t.IsHidden = false;
        if (GetAnn(t, prefix + "Hide") != null) t.IsHidden = true;
        foreach (var c in t.Columns)
        {
            if (GetAnn(c, prefix + "Unhide") != null) c.IsHidden = false;
            if (GetAnn(c, prefix + "Hide") != null) c.IsHidden = true;
        }
        foreach (var m in t.Measures)
        {
            var expression = GetAnn(m, prefix + "Expression");
            if (expression != null) m.Expression = expression;
            if (GetAnn(m, prefix + "Unhide") != null) m.IsHidden = false;
            if (GetAnn(m, prefix + "Hide") != null) m.IsHidden = true;          // original script set false here (bug)
            if (GetAnn(m, prefix + "RemoveKPI") != null) m.KPI = null;
        }

        var dataSourceName = GetAnn(t, prefix + "UpdateDataSource");
        var queryUpdate = GetAnn(t, prefix + "UpdatePartitionQuery");
        foreach (var part in t.Partitions.ToList())
        {
            if (GetAnn(part, prefix + "Remove") != null)
            {
                if (t.Partitions.Count > 1) t.Partitions.Remove(part);
                else log.Add("Partition '" + part.Name + "' in '" + t.Name + "' not removed - a table needs at least one partition.");
                continue;
            }
            var querySource = part.Source as TOM.QueryPartitionSource;
            var mSource = part.Source as TOM.MPartitionSource;
            if (dataSourceName != null && querySource != null)
            {
                var ds = mdl.DataSources.Find(dataSourceName);
                if (ds != null) querySource.DataSource = ds; else log.Add("Data source '" + dataSourceName + "' not found for table '" + t.Name + "'.");
            }
            if (queryUpdate != null)
            {
                var text = querySource != null ? querySource.Query : mSource != null ? mSource.Expression : null;
                if (text != null)
                {
                    var pair = queryUpdate.Split('|');
                    if (pair.Length == 2 && pair[0].Length > 0) text = text.Replace(pair[0], pair[1]);
                    foreach (var ann in part.Annotations) text = text.Replace("%" + ann.Name + "%", ann.Value);
                    if (querySource != null) querySource.Query = text; else mSource.Expression = text;
                }
            }
        }
    }
    foreach (var ds in mdl.DataSources.ToList()) if (GetAnn(ds, prefix + "Remove") != null) mdl.DataSources.Remove(ds);
    foreach (var role in mdl.Roles.ToList())
    {
        if (GetAnn(role, prefix + "Remove") != null) { mdl.Roles.Remove(role); continue; }
        foreach (var t in mdl.Tables)
        {
            var rls = GetAnn(role, prefix + "RLS_" + t.Name);
            if (rls == null) continue;
            var tp = role.TablePermissions.Find(t.Name);
            if (tp == null) { tp = new TOM.TablePermission { Table = t }; role.TablePermissions.Add(tp); }
            tp.FilterExpression = rls;
        }
        var members = GetAnn(role, prefix + "UpdateRoleMembers");
        if (members != null)
        {
            role.Members.Clear();
            foreach (var member in members.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => x.Length > 0))
                role.Members.Add(new TOM.ExternalModelRoleMember { MemberName = member, IdentityProvider = "AzureAD", MemberType = TOM.RoleMemberType.Auto });
        }
        var permission = GetAnn(role, prefix + "ModelPermission");
        if (permission != null) role.ModelPermission = permission == "Administrator" ? TOM.ModelPermission.Administrator : TOM.ModelPermission.Read;
    }

    // Translations of removed objects
    foreach (var culture in mdl.Cultures)
        foreach (var translation in culture.ObjectTranslations.ToList())
            if (translation.Object == null || !IsAttached(translation.Object, mdl)) culture.ObjectTranslations.Remove(translation);

    // Remove the per-perspective annotations of ALL perspectives from the output
    foreach (var obj in AllObjects(mdl))
        foreach (var ann in AnnotationList(obj))
            if (perspectiveNames.Any(pn => ann.Name.StartsWith(pn + "_", StringComparison.OrdinalIgnoreCase))) RemoveAnn(obj, ann);

    var result = RoundTrip(db);
    SetAnn(result.Model, AnnSource, masterName);
    SetAnn(result.Model, AnnPerspective, perspective);
    return result;
}

// ---------------------------------------------------------------------------
// Comparable units (per object) of a TOM model
// ---------------------------------------------------------------------------
Dictionary<string, Unit> CollectUnits(TOM.Model mdl)
{
    var units = new Dictionary<string, Unit>(StringComparer.OrdinalIgnoreCase);
    Action<string, string, string, TOM.MetadataObject, Action<JObject>> add = (kind, table, name, obj, tweak) =>
    {
        var key = kind + "|" + table + "|" + name;
        units[key] = new Unit { Key = key, Kind = kind, Table = table, Name = name, Obj = obj, Hash = HashOf(obj, tweak) };
    };
    foreach (var t in mdl.Tables)
    {
        // Table level: own properties only; children are separate units
        add("Table", "", t.Name, t, j => { foreach (var child in new[] { "columns", "measures", "hierarchies", "partitions", "calculationGroup" }) j.Remove(child); });
        foreach (var c in t.Columns.Where(c => c.Type != TOM.ColumnType.RowNumber)) add("Column", t.Name, c.Name, c, null);
        foreach (var m in t.Measures) add("Measure", t.Name, m.Name, m, null);
        foreach (var h in t.Hierarchies) add("Hierarchy", t.Name, h.Name, h, null);
        foreach (var p in t.Partitions) add("Partition", t.Name, p.Name, p, null);
        if (t.CalculationGroup != null)
            foreach (var ci in t.CalculationGroup.CalculationItems) add("CalculationItem", t.Name, ci.Name, ci, null);
    }
    foreach (var r in mdl.Relationships.OfType<TOM.SingleColumnRelationship>())
        add("Relationship", "", RelationshipKey(r), r, j => j.Remove("name"));
    foreach (var role in mdl.Roles) add("Role", "", role.Name, role, null);
    foreach (var e in mdl.Expressions) add("Expression", "", e.Name, e, null);
    return units;
}

Dictionary<string, string> ReadBaseline(TOM.Model mdl)
{
    var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    var json = GetAnn(mdl, AnnBaseline);
    if (string.IsNullOrWhiteSpace(json)) return map;
    try { foreach (var p in JObject.Parse(json).Properties()) map[p.Name] = (string)p.Value; } catch { }
    return map;
}
void WriteBaseline(TOM.Model target, Dictionary<string, Unit> masterUnits)
{
    var json = new JObject();
    foreach (var u in masterUnits.Values.OrderBy(u => u.Key)) json[u.Key] = u.Hash;
    SetAnn(target, AnnBaseline, json.ToString(Newtonsoft.Json.Formatting.None));
    SetAnn(target, AnnUpdated, DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
}

// ---------------------------------------------------------------------------
// Three-way classification
// ---------------------------------------------------------------------------
List<MergeItem> Classify(Dictionary<string, Unit> master, Dictionary<string, Unit> local, Dictionary<string, string> baseline)
{
    var items = new List<MergeItem>();
    // Children of tables that are added as a whole come with the table
    var newTables = new HashSet<string>(master.Values.Where(u => u.Kind == "Table" && !local.ContainsKey(u.Key)).Select(u => u.Name), StringComparer.OrdinalIgnoreCase);

    foreach (var g in master.Values)
    {
        if (g.Table != "" && newTables.Contains(g.Table)) continue;
        Unit d;
        local.TryGetValue(g.Key, out d);
        string b;
        baseline.TryGetValue(g.Key, out b);
        var item = new MergeItem { Key = g.Key, Kind = g.Kind, Label = g.Label, Master = g, Local = d };

        if (d == null)
        {
            if (b == null) { item.Change = "New in master"; item.Options = new[] { "Add", "Skip" }; item.Action = "Add"; }
            else if (b == g.Hash) { item.Change = "Deleted locally"; item.Options = new[] { "Keep deleted", "Add" }; item.Action = "Keep deleted"; }
            else { item.Change = "Deleted locally, changed in master"; item.Options = new[] { "Keep deleted", "Add" }; }
        }
        else if (d.Hash == g.Hash) continue;
        else if (b != null && b == d.Hash) { item.Change = "Changed in master"; item.Options = new[] { "Take master", "Keep mine" }; item.Action = "Take master"; }
        else if (b != null && b == g.Hash) { item.Change = "Changed locally"; item.Options = new[] { "Keep mine", "Take master" }; item.Action = "Keep mine"; }
        else { item.Change = b == null ? "Different (no fingerprint)" : "Changed in both"; item.Options = new[] { "Take master", "Keep mine" }; }
        items.Add(item);
    }

    foreach (var d in local.Values.Where(u => !master.ContainsKey(u.Key)))
    {
        string b;
        if (!baseline.TryGetValue(d.Key, out b)) continue;      // local-only object: never touched
        items.Add(new MergeItem {
            Key = d.Key, Kind = d.Kind, Label = d.Label, Local = d,
            Change = b == d.Hash ? "Removed from master" : "Removed from master, changed locally",
            Options = new[] { "Keep mine", "Remove" },
            Action = b == d.Hash ? null : "Keep mine"
        });
    }

    return items
        .OrderBy(i => i.Action == null ? 0 : i.Change.StartsWith("Changed locally") || i.Change == "Deleted locally" ? 2 : 1)
        .ThenBy(i => i.Kind == "Table" ? 0 : 1).ThenBy(i => i.Label)
        .ToList();
}

// ---------------------------------------------------------------------------
// Apply chosen actions to the local (derived) TOM model
// ---------------------------------------------------------------------------
void ApplyMerge(List<MergeItem> items, TOM.Model target, List<string> log)
{
    Func<string, TOM.Table> tableOf = name => target.Tables.Find(name);

    // 1. Removals (relationships and hierarchies before columns, children before tables)
    var removals = items.Where(i => i.Action == "Remove" || (i.Action == "Take master" && (i.Kind == "Relationship" || i.Kind == "Hierarchy" || i.Kind == "Role"))).ToList();
    foreach (var i in removals.OrderBy(i => i.Kind == "Relationship" ? 0 : i.Kind == "Hierarchy" ? 1 : i.Kind == "Table" ? 9 : 5))
    {
        var obj = i.Local == null ? null : i.Local.Obj;
        if (obj == null || !IsAttached(obj, target)) continue;
        if (obj is TOM.Relationship) target.Relationships.Remove((TOM.Relationship)obj);
        else if (obj is TOM.Hierarchy) ((TOM.Hierarchy)obj).Table.Hierarchies.Remove((TOM.Hierarchy)obj);
        else if (obj is TOM.ModelRole) target.Roles.Remove((TOM.ModelRole)obj);
        else if (obj is TOM.Measure) ((TOM.Measure)obj).Table.Measures.Remove((TOM.Measure)obj);
        else if (obj is TOM.CalculationItem) ((TOM.CalculationItem)obj).CalculationGroup.CalculationItems.Remove((TOM.CalculationItem)obj);
        else if (obj is TOM.NamedExpression) target.Expressions.Remove((TOM.NamedExpression)obj);
        else if (obj is TOM.Partition)
        {
            var p = (TOM.Partition)obj;
            if (p.Table.Partitions.Count > 1) p.Table.Partitions.Remove(p); else log.Add("Kept partition " + i.Label + " - a table needs at least one partition.");
        }
        else if (obj is TOM.Column)
        {
            var c = (TOM.Column)obj;
            foreach (var r in target.Relationships.OfType<TOM.SingleColumnRelationship>().Where(r => r.FromColumn == c || r.ToColumn == c).ToList()) target.Relationships.Remove(r);
            foreach (var h in c.Table.Hierarchies.Where(h => h.Levels.Any(l => l.Column == c)).ToList()) c.Table.Hierarchies.Remove(h);
            foreach (var other in c.Table.Columns.Where(o => o.SortByColumn == c).ToList()) other.SortByColumn = null;
            c.Table.Columns.Remove(c);
        }
        else if (obj is TOM.Table)
        {
            var t = (TOM.Table)obj;
            foreach (var r in target.Relationships.Where(r => r.FromTable == t || r.ToTable == t).ToList()) target.Relationships.Remove(r);
            foreach (var role in target.Roles)
                foreach (var tp in role.TablePermissions.Where(tp => tp.Table == t).ToList()) role.TablePermissions.Remove(tp);
            target.Tables.Remove(t);
        }
    }

    var adds = items.Where(i => i.Action == "Add" || i.Action == "Take master").ToList();

    // 2. Tables (whole table when new; own properties when changed)
    foreach (var i in adds.Where(i => i.Kind == "Table"))
    {
        var source = (TOM.Table)i.Master.Obj;
        var existing = tableOf(source.Name);
        if (existing == null)
        {
            var clone = (TOM.Table)CloneObject(source);
            foreach (var c in clone.Columns) c.Variations.Clear();
            target.Tables.Add(clone);
        }
        else
        {
            existing.Description = source.Description;
            existing.IsHidden = source.IsHidden;
            existing.DataCategory = source.DataCategory;
            existing.IsPrivate = source.IsPrivate;
        }
    }

    // 3. Columns (update in place so relationships/hierarchies stay attached)
    var skipProps = new HashSet<string> { "Name", "Type", "Parent", "Table", "Model", "SortByColumn", "LineageTag", "SourceLineageTag", "ModifiedTime", "StructureModifiedTime", "RefreshedTime", "State", "ErrorMessage", "ObjectType", "IsRemoved", "IsAvailableInMDX" };
    foreach (var i in adds.Where(i => i.Kind == "Column"))
    {
        var source = (TOM.Column)i.Master.Obj;
        var table = tableOf(i.Master.Table);
        if (table == null) { log.Add("Skipped column " + i.Label + " - table missing locally."); continue; }
        var existing = table.Columns.Find(source.Name);
        if (existing == null)
        {
            var clone = (TOM.Column)CloneObject(source);
            clone.SortByColumn = null;
            clone.Variations.Clear();
            table.Columns.Add(clone);
            existing = clone;
        }
        else
        {
            foreach (var prop in source.GetType().GetProperties().Where(p => p.CanRead && p.CanWrite && !skipProps.Contains(p.Name)
                && (p.PropertyType.IsPrimitive || p.PropertyType.IsEnum || p.PropertyType == typeof(string) || p.PropertyType == typeof(double))))
            {
                try { prop.SetValue(existing, prop.GetValue(source, null), null); } catch { }
            }
        }
        existing.SortByColumn = source.SortByColumn == null ? null : table.Columns.Find(source.SortByColumn.Name);
    }

    // 4. Measures, calculation items, partitions, expressions (replace object)
    foreach (var i in adds.Where(i => i.Kind == "Measure" || i.Kind == "CalculationItem" || i.Kind == "Partition"))
    {
        var table = tableOf(i.Master.Table);
        if (table == null) { log.Add("Skipped " + i.Label + " - table missing locally."); continue; }
        if (i.Kind == "Measure")
        {
            var old = table.Measures.Find(i.Master.Name);
            if (old != null) table.Measures.Remove(old);
            table.Measures.Add((TOM.Measure)CloneObject(i.Master.Obj));
        }
        else if (i.Kind == "CalculationItem")
        {
            if (table.CalculationGroup == null) { log.Add("Skipped " + i.Label + " - not a calculation group locally."); continue; }
            var old = table.CalculationGroup.CalculationItems.Find(i.Master.Name);
            if (old != null) table.CalculationGroup.CalculationItems.Remove(old);
            table.CalculationGroup.CalculationItems.Add((TOM.CalculationItem)CloneObject(i.Master.Obj));
        }
        else
        {
            var old = table.Partitions.Find(i.Master.Name);
            var clone = (TOM.Partition)CloneObject(i.Master.Obj);
            var legacySource = clone.Source as TOM.QueryPartitionSource;
            if (legacySource != null && ((TOM.QueryPartitionSource)((TOM.Partition)i.Master.Obj).Source).DataSource != null)
                legacySource.DataSource = target.DataSources.Find(((TOM.QueryPartitionSource)((TOM.Partition)i.Master.Obj).Source).DataSource.Name);
            if (old != null) table.Partitions.Remove(old);
            table.Partitions.Add(clone);
        }
    }
    foreach (var i in adds.Where(i => i.Kind == "Expression"))
    {
        var old = target.Expressions.Find(i.Master.Name);
        if (old != null) target.Expressions.Remove(old);
        target.Expressions.Add((TOM.NamedExpression)CloneObject(i.Master.Obj));
    }

    // 5. Hierarchies (rebuilt on local columns; old one removed in step 1)
    foreach (var i in adds.Where(i => i.Kind == "Hierarchy"))
    {
        var source = (TOM.Hierarchy)i.Master.Obj;
        var table = tableOf(i.Master.Table);
        if (table == null || table.Hierarchies.Find(source.Name) != null) continue;
        var h = new TOM.Hierarchy { Name = source.Name, Description = source.Description, DisplayFolder = source.DisplayFolder, IsHidden = source.IsHidden, LineageTag = source.LineageTag };
        bool ok = true;
        foreach (var level in source.Levels.OrderBy(l => l.Ordinal))
        {
            var column = table.Columns.Find(level.Column.Name);
            if (column == null) { ok = false; break; }
            h.Levels.Add(new TOM.Level { Name = level.Name, Ordinal = level.Ordinal, Column = column, LineageTag = level.LineageTag });
        }
        if (ok) table.Hierarchies.Add(h); else log.Add("Skipped hierarchy " + i.Label + " - a level column is missing locally.");
    }

    // 6. Relationships (rebuilt on local columns; old one removed in step 1)
    foreach (var i in adds.Where(i => i.Kind == "Relationship"))
    {
        var source = (TOM.SingleColumnRelationship)i.Master.Obj;
        var fromTable = tableOf(source.FromTable.Name);
        var toTable = tableOf(source.ToTable.Name);
        var fromColumn = fromTable == null ? null : fromTable.Columns.Find(source.FromColumn.Name);
        var toColumn = toTable == null ? null : toTable.Columns.Find(source.ToColumn.Name);
        if (fromColumn == null || toColumn == null) { log.Add("Skipped relationship " + i.Label + " - a column is missing locally."); continue; }
        target.Relationships.Add(new TOM.SingleColumnRelationship {
            Name = Guid.NewGuid().ToString(),
            FromColumn = fromColumn, ToColumn = toColumn,
            FromCardinality = source.FromCardinality, ToCardinality = source.ToCardinality,
            IsActive = source.IsActive,
            CrossFilteringBehavior = source.CrossFilteringBehavior,
            SecurityFilteringBehavior = source.SecurityFilteringBehavior,
            JoinOnDateBehavior = source.JoinOnDateBehavior,
            RelyOnReferentialIntegrity = source.RelyOnReferentialIntegrity
        });
    }

    // 7. Roles (rebuilt on local tables; old one removed in step 1)
    foreach (var i in adds.Where(i => i.Kind == "Role"))
    {
        var source = (TOM.ModelRole)i.Master.Obj;
        if (target.Roles.Find(source.Name) != null) continue;
        var role = new TOM.ModelRole { Name = source.Name, Description = source.Description, ModelPermission = source.ModelPermission };
        foreach (var member in source.Members) role.Members.Add((TOM.ModelRoleMember)CloneObject(member));
        foreach (var tp in source.TablePermissions)
        {
            var table = tableOf(tp.Table.Name);
            if (table == null) continue;
            var newTp = new TOM.TablePermission { Table = table, FilterExpression = tp.FilterExpression, MetadataPermission = tp.MetadataPermission };
            foreach (var cp in tp.ColumnPermissions)
            {
                var column = table.Columns.Find(cp.Column.Name);
                if (column != null) newTp.ColumnPermissions.Add(new TOM.ColumnPermission { Column = column, MetadataPermission = cp.MetadataPermission });
            }
            role.TablePermissions.Add(newTp);
        }
        target.Roles.Add(role);
    }
}

// ---------------------------------------------------------------------------
// Disk I/O
// ---------------------------------------------------------------------------
string DefinitionFolder(string semanticModelFolder) { return Path.Combine(semanticModelFolder, "definition"); }
bool IsBim(string semanticModelFolder) { return !Directory.Exists(DefinitionFolder(semanticModelFolder)) && File.Exists(Path.Combine(semanticModelFolder, "model.bim")); }

TOM.Database LoadModel(string semanticModelFolder)
{
    return IsBim(semanticModelFolder)
        ? TOM.JsonSerializer.DeserializeDatabase(File.ReadAllText(Path.Combine(semanticModelFolder, "model.bim")))
        : TOM.TmdlSerializer.DeserializeDatabaseFromFolder(DefinitionFolder(semanticModelFolder));
}

string SaveModel(TOM.Database db, string semanticModelFolder)
{
    // Backup of the current definition, then replace it
    string backup = null;
    var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
    var backupRoot = Path.Combine(Path.GetTempPath(), "MasterModelBackup", Path.GetFileName(semanticModelFolder) + "_" + stamp);
    if (IsBim(semanticModelFolder))
    {
        var bim = Path.Combine(semanticModelFolder, "model.bim");
        Directory.CreateDirectory(backupRoot);
        File.Copy(bim, Path.Combine(backupRoot, "model.bim"));
        backup = backupRoot;
        File.WriteAllText(bim, TOM.JsonSerializer.SerializeDatabase(db));
        return backup;
    }
    var definition = DefinitionFolder(semanticModelFolder);
    var staging = definition + "_new_" + stamp;
    TOM.TmdlSerializer.SerializeDatabaseToFolder(db, staging);
    if (Directory.Exists(definition))
    {
        // Copy (not move): %TEMP% may be on another drive than the repository
        CopyDirectory(definition, backupRoot);
        backup = backupRoot;
        Directory.Delete(definition, true);
    }
    Directory.Move(staging, definition);
    return backup;
}

void CreateModelFolder(TOM.Database db, string semanticModelFolder, string name)
{
    Directory.CreateDirectory(semanticModelFolder);
    TOM.TmdlSerializer.SerializeDatabaseToFolder(db, DefinitionFolder(semanticModelFolder));
    File.WriteAllText(Path.Combine(semanticModelFolder, "definition.pbism"),
        new JObject {
            ["$schema"] = "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
            ["version"] = "4.2",
            ["settings"] = new JObject()
        }.ToString());
    File.WriteAllText(Path.Combine(semanticModelFolder, ".platform"),
        new JObject {
            ["$schema"] = "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            ["metadata"] = new JObject { ["type"] = "SemanticModel", ["displayName"] = name },
            ["config"] = new JObject { ["version"] = "2.0", ["logicalId"] = Guid.NewGuid().ToString() }
        }.ToString());
}

// Existing models created by this macro: read the annotations without loading the model
List<DerivedModel> FindDerivedModels(string folder)
{
    var found = new List<DerivedModel>();
    foreach (var dir in Directory.GetDirectories(folder, "*.SemanticModel"))
    {
        var model = new DerivedModel { Folder = dir, Name = Path.GetFileName(dir).Substring(0, Path.GetFileName(dir).Length - ".SemanticModel".Length) };
        try
        {
            var modelTmdl = Path.Combine(dir, "definition", "model.tmdl");
            var bim = Path.Combine(dir, "model.bim");
            string text = File.Exists(modelTmdl) ? File.ReadAllText(modelTmdl) : File.Exists(bim) ? File.ReadAllText(bim) : "";
            var p = Regex.Match(text, "annotation\\s+" + AnnPerspective + "\\s*=\\s*(.+)");
            var s = Regex.Match(text, "annotation\\s+" + AnnSource + "\\s*=\\s*(.+)");
            if (!p.Success) p = Regex.Match(text, "\"name\"\\s*:\\s*\"" + AnnPerspective + "\"\\s*,\\s*\"value\"\\s*:\\s*\"([^\"]*)\"");
            if (!s.Success) s = Regex.Match(text, "\"name\"\\s*:\\s*\"" + AnnSource + "\"\\s*,\\s*\"value\"\\s*:\\s*\"([^\"]*)\"");
            if (p.Success) model.Perspective = p.Groups[1].Value.Trim().Trim('"', '\'');
            if (s.Success) model.Source = s.Groups[1].Value.Trim().Trim('"', '\'');
            var perspectivesFolder = Path.Combine(dir, "definition", "perspectives");
            model.HasPerspectives = (Directory.Exists(perspectivesFolder) && Directory.GetFiles(perspectivesFolder, "*.tmdl").Length > 0)
                || (File.Exists(bim) && Regex.IsMatch(text, "\"perspectives\"\\s*:\\s*\\[\\s*\\{"));
        }
        catch { }
        found.Add(model);
    }
    return found;
}

// ---------------------------------------------------------------------------
// UI helpers (same look as the other macros)
// ---------------------------------------------------------------------------
var editBack = Color.FromArgb(255, 244, 179);
var editHeaderBack = Color.FromArgb(255, 214, 102);
var decisionBack = Color.FromArgb(255, 222, 222);

Size DialogSize(double w, double h)
{
    var wa = Screen.PrimaryScreen.WorkingArea;
    return new Size((int)(wa.Width * w), (int)(wa.Height * h));
}
TableLayoutPanel CreateDialogLayout(Form dlg, string title, int rows, Size size)
{
    dlg.Text = title;
    dlg.Size = size;
    dlg.MinimumSize = new Size(700, 440);
    dlg.StartPosition = FormStartPosition.CenterScreen;
    dlg.FormBorderStyle = FormBorderStyle.Sizable;
    dlg.MaximizeBox = true;
    dlg.MinimizeBox = false;
    dlg.Font = new Font("Segoe UI", 9);
    var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = rows, ColumnCount = 1, Padding = new Padding(12) };
    dlg.Controls.Add(panel);
    return panel;
}
Label Heading(string text)
{
    return new Label { Text = text, AutoSize = true, MaximumSize = new Size(1200, 0), Font = new Font("Segoe UI", 10, FontStyle.Bold), Margin = new Padding(0, 0, 0, 8) };
}
Label Hint(string text)
{
    return new Label { Text = text, AutoSize = true, MaximumSize = new Size(1200, 0), ForeColor = Color.DimGray, Margin = new Padding(0, 6, 0, 0) };
}
DataGridView EditGrid()
{
    var grid = new DataGridView {
        Dock = DockStyle.Fill, AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.CellSelect, MultiSelect = false,
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells, BackgroundColor = SystemColors.Window,
        EnableHeadersVisualStyles = false, EditMode = DataGridViewEditMode.EditOnEnter
    };
    grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
    grid.DefaultCellStyle.ForeColor = Color.FromArgb(90, 90, 90);
    grid.CurrentCellDirtyStateChanged += (s, e) =>
    {
        if (grid.IsCurrentCellDirty && (grid.CurrentCell is DataGridViewCheckBoxCell || grid.CurrentCell is DataGridViewComboBoxCell))
            grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
    };
    return grid;
}
void Highlight(DataGridViewColumn column)
{
    column.DefaultCellStyle = new DataGridViewCellStyle {
        BackColor = editBack, ForeColor = Color.Black, Font = new Font("Segoe UI", 9, FontStyle.Bold),
        SelectionBackColor = editHeaderBack, SelectionForeColor = Color.Black
    };
    column.HeaderCell.Style.BackColor = editHeaderBack;
    column.HeaderCell.Style.ForeColor = Color.Black;
}

// ---------------------------------------------------------------------------
// Step 1: checks and output folder
// ---------------------------------------------------------------------------
if (perspectiveNames.Count == 0) { Error("The open model has no perspectives. Create a perspective per target model first."); return; }

var settingsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TabularEditor3", "MasterModelMacro.lastfolder.txt");
string outputFolder = null;
using (var fbd = new FolderBrowserDialog())
{
    fbd.Description = "Select the folder that contains the master .SemanticModel folder. New perspective models are created here, existing ones are found here.";
    try { if (File.Exists(settingsFile)) fbd.SelectedPath = File.ReadAllText(settingsFile).Trim(); } catch { }
    if (fbd.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
    outputFolder = fbd.SelectedPath;
}
try { Directory.CreateDirectory(Path.GetDirectoryName(settingsFile)); File.WriteAllText(settingsFile, outputFolder); } catch { }

var derivedModels = FindDerivedModels(outputFolder);

// PBIP models have no database name: use the master's folder name (the model in this
// folder that still has perspectives) for default names and the MasterModel_Source stamp
if (masterName == "Model")
{
    var masterCandidates = derivedModels.Where(d => d.Perspective == null && d.HasPerspectives).ToList();
    if (masterCandidates.Count == 1) masterName = masterCandidates[0].Name;
}

// ---------------------------------------------------------------------------
// Step 2: perspectives grid (multi-select, editable names)
// ---------------------------------------------------------------------------
var perspectiveRows = perspectiveNames.Select(pn => new PerspectiveRow {
    Perspective = pn,
    Tables = Model.Tables.Count(t => t.InPerspective[pn]),
    Columns = Model.AllColumns.Count(c => c.InPerspective[pn]),
    Measures = Model.AllMeasures.Count(m => m.InPerspective[pn])
}).ToList();

// name -> (Create/Update, status text)
Func<string, string, Tuple<string, string, bool>> Evaluate = (perspective, name) =>
{
    name = (name ?? "").Trim();
    if (name.Length == 0) return Tuple.Create("-", "Enter a model name", false);
    if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return Tuple.Create("-", "Name contains characters not allowed in a folder name", false);
    var existing = derivedModels.FirstOrDefault(d => SameName(d.Name, name));
    if (existing == null) return Tuple.Create("Create", "New: " + name + ".SemanticModel", true);
    if (existing.Perspective == null && existing.HasPerspectives) return Tuple.Create("-", "Exists and has perspectives (looks like a master model) - choose another name", false);
    if (existing.Perspective == null) return Tuple.Create("Update", "Exists, not created by this macro - every difference will be asked", true);
    if (!SameName(existing.Perspective, perspective)) return Tuple.Create("-", "Exists for perspective '" + existing.Perspective + "' - choose another name", false);
    return Tuple.Create("Update", "Exists - update from master" + (existing.Source != null && !SameName(existing.Source, masterName) ? " (created from '" + existing.Source + "')" : ""), true);
};

var runList = new List<Tuple<string, string, string>>();   // perspective, name, action
using (var dlg = new Form())
{
    var panel = CreateDialogLayout(dlg, "Master model - perspective models", 4, DialogSize(0.75, 0.6));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var grid = EditGrid();
    var includeColumn = new DataGridViewCheckBoxColumn { Name = "Include", HeaderText = "Include", Width = 60 };
    Highlight(includeColumn);
    grid.Columns.Add(includeColumn);
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Perspective", HeaderText = "Perspective", Width = 180, ReadOnly = true });
    var nameColumn = new DataGridViewTextBoxColumn { Name = "ModelName", HeaderText = "Model name (edit)", Width = 260 };
    Highlight(nameColumn);
    grid.Columns.Add(nameColumn);
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Action", HeaderText = "Action", Width = 70, ReadOnly = true });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "Status", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 220 });
    grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Content", HeaderText = "In perspective", Width = 190, ReadOnly = true });

    foreach (var pr in perspectiveRows)
    {
        var existing = derivedModels.FirstOrDefault(d => d.Perspective != null && SameName(d.Perspective, pr.Perspective));
        var defaultName = existing != null ? existing.Name : pr.Perspective + " " + masterName;
        int index = grid.Rows.Add(false, pr.Perspective, defaultName, "", "", pr.Tables + " tables, " + pr.Columns + " columns, " + pr.Measures + " measures");
        grid.Rows[index].Tag = pr;
    }

    Action<int> refresh = rowIndex =>
    {
        var row = grid.Rows[rowIndex];
        var result = Evaluate((string)row.Cells["Perspective"].Value, Convert.ToString(row.Cells["ModelName"].Value));
        row.Cells["Action"].Value = result.Item1;
        row.Cells["Status"].Value = result.Item2;
        row.Cells["Status"].Style.ForeColor = !result.Item3 ? Color.DarkRed : result.Item1 == "Update" ? Color.DarkOrange : Color.DarkGreen;
    };
    for (int i = 0; i < grid.Rows.Count; i++) refresh(i);
    grid.CellValueChanged += (s, e) => { if (e.RowIndex >= 0) refresh(e.RowIndex); };

    var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
    var okButton = new Button { Text = "Continue", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 0, 8, 0) };
    var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
    buttons.Controls.AddRange(new Control[] { okButton, cancelButton });
    okButton.Click += (s, e) =>
    {
        grid.EndEdit();
        runList.Clear();
        foreach (DataGridViewRow row in grid.Rows)
        {
            if (!Convert.ToBoolean(row.Cells["Include"].Value)) continue;
            var perspective = (string)row.Cells["Perspective"].Value;
            var name = Convert.ToString(row.Cells["ModelName"].Value).Trim();
            var result = Evaluate(perspective, name);
            if (!result.Item3)
            {
                MessageBox.Show(perspective + ": " + result.Item2, "Model name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                grid.CurrentCell = row.Cells["ModelName"];
                return;
            }
            if (runList.Any(r => SameName(r.Item2, name)))
            {
                MessageBox.Show("The name '" + name + "' is used twice.", "Model name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            runList.Add(Tuple.Create(perspective, name, result.Item1));
        }
        if (runList.Count == 0)
        {
            MessageBox.Show("Tick at least one perspective.", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        dlg.DialogResult = DialogResult.OK;
        dlg.Close();
    };

    panel.Controls.Add(Heading("Master: " + masterName + "  |  Folder: " + outputFolder + "\nTick the perspectives to process. The model name decides the action: a new name creates a model, an existing name updates it."), 0, 0);
    panel.Controls.Add(grid, 0, 1);
    panel.Controls.Add(Hint("Create writes <Name>.SemanticModel (TMDL, no report) in the folder above. Update merges master changes into the existing model and asks when both sides changed. The master model is not changed."), 0, 2);
    panel.Controls.Add(buttons, 0, 3);
    dlg.CancelButton = cancelButton;
    if (dlg.ShowDialog() != DialogResult.OK) { Info("Cancelled - no changes made."); return; }
}

// ---------------------------------------------------------------------------
// Step 3: create / update each selected perspective model
// ---------------------------------------------------------------------------
var summary = new System.Data.DataTable();
summary.Columns.Add("Model");
summary.Columns.Add("Perspective");
summary.Columns.Add("Action");
summary.Columns.Add("Result");
summary.Columns.Add("Notes");

foreach (var run in runList)
{
    var perspective = run.Item1;
    var name = run.Item2;
    var semanticModelFolder = Path.Combine(outputFolder, name + ".SemanticModel");
    var log = new List<string>();

    try
    {
        var generated = BuildPerspectiveModel(perspective, name, log);
        var masterUnits = CollectUnits(generated.Model);

        if (run.Item3 == "Create")
        {
            WriteBaseline(generated.Model, masterUnits);
            CreateModelFolder(generated, semanticModelFolder, name);
            summary.Rows.Add(name, perspective, "Create",
                "Created with " + generated.Model.Tables.Count + " tables, " + generated.Model.Tables.Sum(t => t.Measures.Count) + " measures",
                string.Join("; ", log));
            continue;
        }

        // ---- Update: three-way compare ----
        var localDb = LoadModel(semanticModelFolder);
        var baseline = ReadBaseline(localDb.Model);
        var localUnits = CollectUnits(localDb.Model);
        var items = Classify(masterUnits, localUnits, baseline);

        if (items.Count == 0)
        {
            summary.Rows.Add(name, perspective, "Update", "Already up to date - no changes", "");
            continue;
        }

        bool skipModel = false;
        bool cancelAll = false;
        using (var dlg = new Form())
        {
            var panel = CreateDialogLayout(dlg, "Update " + name + " from master", 5, DialogSize(0.9, 0.9));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var grid = EditGrid();
            var actionColumn = new DataGridViewComboBoxColumn { Name = "Action", HeaderText = "Action (edit)", Width = 130, FlatStyle = FlatStyle.Flat };
            actionColumn.Items.AddRange("Choose...", "Add", "Skip", "Take master", "Keep mine", "Keep deleted", "Remove");
            Highlight(actionColumn);
            grid.Columns.Add(actionColumn);
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Change", HeaderText = "Change", Width = 230, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Kind", HeaderText = "Type", Width = 110, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Object", HeaderText = "Object", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 240 });

            foreach (var item in items)
            {
                int index = grid.Rows.Add(item.Action ?? "Choose...", item.Change, item.Kind, item.Label +
                    (item.Kind == "Table" && item.Local == null ? "  (whole table incl. columns and measures)" : ""));
                var row = grid.Rows[index];
                row.Tag = item;
                // Only the options valid for this change are offered
                var cell = (DataGridViewComboBoxCell)row.Cells["Action"];
                cell.Items.Clear();
                if (item.Action == null) cell.Items.Add("Choose...");
                foreach (var option in item.Options) cell.Items.Add(option);
                cell.Value = item.Action ?? "Choose...";
                if (item.Action == null) foreach (DataGridViewCell c in row.Cells) if (c.OwningColumn.Name != "Action") c.Style.BackColor = decisionBack;
            }
            grid.CellValueChanged += (s, e) =>
            {
                if (e.RowIndex < 0 || grid.Columns[e.ColumnIndex].Name != "Action") return;
                var row = grid.Rows[e.RowIndex];
                var chosen = Convert.ToString(row.Cells["Action"].Value) != "Choose...";
                foreach (DataGridViewCell c in row.Cells) if (c.OwningColumn.Name != "Action") c.Style.BackColor = chosen ? SystemColors.Window : decisionBack;
            };

            // Side-by-side definitions of the selected row
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical };
            Func<string, TextBox> codeBox = title => new TextBox {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false,
                Font = new Font("Consolas", 9), BackColor = SystemColors.Window
            };
            var masterBox = codeBox("Master");
            var localBox = codeBox("Local");
            var masterLabel = new Label { Text = "Master (generated from perspective)", Dock = DockStyle.Top, Font = new Font("Segoe UI", 9, FontStyle.Bold), Height = 20 };
            var localLabel = new Label { Text = "Local (" + name + ")", Dock = DockStyle.Top, Font = new Font("Segoe UI", 9, FontStyle.Bold), Height = 20 };
            split.Panel1.Controls.Add(masterBox); split.Panel1.Controls.Add(masterLabel);
            split.Panel2.Controls.Add(localBox); split.Panel2.Controls.Add(localLabel);
            grid.SelectionChanged += (s, e) =>
            {
                if (grid.CurrentRow == null) return;
                var item = (MergeItem)grid.CurrentRow.Tag;
                masterBox.Text = Readable(item.Master == null ? null : item.Master.Obj);
                localBox.Text = Readable(item.Local == null ? null : item.Local.Obj);
            };
            dlg.Shown += (s, e) => { split.SplitterDistance = split.Width / 2; if (grid.Rows.Count > 0) grid.CurrentCell = grid.Rows[0].Cells["Action"]; };

            int decisions = items.Count(i => i.Action == null);
            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 8, 0, 0) };
            var applyButton = new Button { Text = "Apply", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), Margin = new Padding(0, 0, 8, 0) };
            var skipButton = new Button { Text = "Skip this model", AutoSize = true, Margin = new Padding(0, 0, 8, 0) };
            var cancelButton = new Button { Text = "Cancel all", AutoSize = true };
            buttons.Controls.AddRange(new Control[] { applyButton, skipButton, cancelButton });
            applyButton.Click += (s, e) =>
            {
                grid.EndEdit();
                foreach (DataGridViewRow row in grid.Rows)
                {
                    var value = Convert.ToString(row.Cells["Action"].Value);
                    if (value == "Choose...")
                    {
                        MessageBox.Show("Choose an action for every red row (" + row.Cells["Object"].Value + ").", "Decision needed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        grid.CurrentCell = row.Cells["Action"];
                        return;
                    }
                    ((MergeItem)row.Tag).Action = value;
                }
                dlg.DialogResult = DialogResult.OK;
                dlg.Close();
            };
            skipButton.Click += (s, e) => { skipModel = true; dlg.DialogResult = DialogResult.OK; dlg.Close(); };
            cancelButton.Click += (s, e) => { cancelAll = true; dlg.DialogResult = DialogResult.Cancel; dlg.Close(); };

            panel.Controls.Add(Heading("Update '" + name + "' from perspective '" + perspective + "': " + items.Count + " difference(s)" +
                (decisions > 0 ? ", " + decisions + " need your decision (red rows)." : ".") +
                (baseline.Count == 0 ? "\nThis model has no fingerprint from this macro, so every difference is asked." : "")), 0, 0);
            panel.Controls.Add(grid, 0, 1);
            panel.Controls.Add(split, 0, 2);
            panel.Controls.Add(Hint("Add / Take master = use the master version. Keep mine = leave the local version. Keep deleted = do not bring back an object you deleted. Remove = delete an object the master no longer has. " +
                "Objects that only exist locally are never touched. Select a row to compare the definitions."), 0, 3);
            panel.Controls.Add(buttons, 0, 4);
            dlg.CancelButton = cancelButton;
            dlg.ShowDialog();
        }

        if (cancelAll)
        {
            summary.Rows.Add(name, perspective, "Update", "Cancelled", "");
            break;
        }
        if (skipModel)
        {
            summary.Rows.Add(name, perspective, "Update", "Skipped - no changes", "");
            continue;
        }

        ApplyMerge(items, localDb.Model, log);
        // New fingerprint = the master version as of now; local-only objects keep being ignored
        WriteBaseline(localDb.Model, masterUnits);
        SetAnn(localDb.Model, AnnSource, masterName);
        SetAnn(localDb.Model, AnnPerspective, perspective);
        var backup = SaveModel(localDb, semanticModelFolder);

        summary.Rows.Add(name, perspective, "Update",
            items.Count(i => i.Action == "Add") + " added, " + items.Count(i => i.Action == "Take master") + " updated, " +
            items.Count(i => i.Action == "Remove") + " removed, " + items.Count(i => i.Action == "Keep mine") + " kept local",
            (backup != null ? "Backup: " + backup + ". " : "") + string.Join("; ", log));
    }
    catch (Exception ex)
    {
        summary.Rows.Add(name, perspective, run.Item3, "FAILED - nothing written for this model", ex.Message + (ex.InnerException != null ? " | " + ex.InnerException.Message : ""));
    }
}

Output(summary);
Info("Done. See the output grid for each model.\n\n" +
    "The master model was not changed. New/updated models are written in:\n" + outputFolder + "\n\n" +
    "Next: open <Name>.SemanticModel\\definition in Tabular Editor to review, then deploy or sync the folder with Fabric Git integration. " +
    "Review Copilot / linguistic settings for the smaller models.");
