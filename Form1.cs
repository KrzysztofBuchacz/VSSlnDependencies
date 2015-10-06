using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSSlnDependencies
{
  public partial class Form1 : Form
  {
    string slnFileName = "";
    Dictionary<string, ProjectInfo> projects = new Dictionary<string, ProjectInfo>();

    public Form1()
    {
      InitializeComponent();
      toolStripStatusLabel1.Text = "No .sln loaded";
    }

    private void toolStripButton1_Click(object sender, EventArgs e)
    {
      try
      {
        OpenFileDialog openExtensFile = new OpenFileDialog();
        openExtensFile.Filter = "Solution Files (*.sln)|*.sln|All files (*.*)|*.*";
        if (openExtensFile.ShowDialog() == DialogResult.OK)
        {
          slnFileName = openExtensFile.FileName;
          LoadSln();
          toolStripButton4.Enabled = true;
          toolStripButton5.Enabled = true;
          toolStripButton6.Enabled = true;
        }
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
      }
    }

    void LoadSln()
    {
      int currentLine = 0;
      string projName = "";
      try
      {
        projects.Clear();
        listView1.Items.Clear();
        FileInfo slnFile = new FileInfo(slnFileName);
        StreamReader sr = File.OpenText(slnFileName);
        bool inProject = false;
        bool inProjectDependencies = false;
        while (sr.Peek() >= 0)
        {
          currentLine++;
          string line = sr.ReadLine();
          if (inProject)
          {
            line = line.Trim();
            if (inProjectDependencies)
            {
              if (line == "EndProjectSection")
              {
                inProjectDependencies = false;
              }
              else
              {
                string guid = line.Split('=')[0];
                guid = guid.Trim(new char[] { '\\', '"', ' ' });
                projects[projName].dependencies.Add(new Guid(guid));
              }
            }
            if (line == "ProjectSection(ProjectDependencies) = postProject")
            {
              inProjectDependencies = true;
            }
            else if (line == "EndProject")
            {
              inProject = false;
            }
          }
          else
          {
            if (line.Contains("Project(\"{2150E333-8FDC-42A3-9474-1A3956D46DE8"))
            {
              inProject = true;
            }
            else if (line.Contains("Project(\"{"))
            {
              inProject = true;
              string[] names = line.Split('=')[1].Split(',');
              projName = names[0];
              string projPath = names[1];
              string projGuid = names[2];

              projPath = projPath.Trim(new char[] { '\\', '"', ' ' });
              projGuid = projGuid.Trim(new char[] { '\\', '"', ' ' });

              ProjectInfo projInfo = new ProjectInfo();
              projInfo.dependencies = new List<Guid>();
              projInfo.references = new List<Guid>();
              projInfo.guid = new Guid(projGuid);
              if (File.Exists(projPath))
                projInfo.fileInfo = new FileInfo(projPath);
              else
                projInfo.fileInfo = new FileInfo(slnFile.DirectoryName + "\\" + projPath);
              projInfo.binName = FindBinName(projInfo.fileInfo);

              projects.Add(projName, projInfo);
            }
          }
        }
        sr.Close();
        currentLine = 0;
        // find references
        foreach (ProjectInfo pi in projects.Values)
        {
          sr = File.OpenText(pi.fileInfo.FullName);
          //System.Diagnostics.Debug.Assert(!pi.fileInfo.FullName.Contains("ROITreeCtrl"), "ROITreeCtrl");
          bool inReference = false;
          while (sr.Peek() >= 0)
          {
            string line = sr.ReadLine();
            if (inReference)
            {
              if (line.Contains("<\\Reference "))
              {
                inReference = false;
              }
              else
              {
                if (line.Contains("<HintPath>"))
                {
                  string refpath = line.Substring(line.IndexOf(">") + 1, line.IndexOf("</") - line.IndexOf(">") - 1);
                  string name = refpath.Split(new char[] { '\\', '/' }).Last();
                  name = name.Replace(".exe", "");
                  name = name.Replace(".dll", "");
                  name = name.Replace(".lib", "");
                  name = name.ToLower();
                  KeyValuePair<string, ProjectInfo> found = projects.Where(p => p.Value.binName == name).FirstOrDefault();
                  if (found.Key != null)
                  {
                    pi.references.Add(found.Value.guid);
                  }
                }
              }
            }
            else
            {
              if (line.Contains("<Reference "))
              {
                inReference = true;
              }
              else if (line.Contains("<AdditionalDependencies>"))
              {
                if (line.Contains("</AdditionalDependencies>"))
                {
                  string[] deps = line.Substring(line.IndexOf(">") + 1, line.IndexOf("</") - line.IndexOf(">") - 1).Split(';');
                  foreach (string dep in deps)
                  {
                    string name = dep.Split(new char[] { '\\', '/' }).Last();
                    name = name.ToLower();
                    name = name.Replace("%(additionaldependencies)", "");
                    name = name.Replace("$(outdir)", "");
                    name = name.Replace(".lib", "");
                    KeyValuePair<string, ProjectInfo> found = projects.Where(p => p.Value.binName == name).FirstOrDefault();
                    if (found.Key != null)
                    {
                      pi.references.Add(found.Value.guid);
                    }
                  }
                }
                else
                {
                  System.Diagnostics.Debug.Assert(false, "<AdditionalDependencies>");
                }
              }
            }
          }
          sr.Close();
        }
        foreach (ProjectInfo pi in projects.Values)
        {
          foreach (Guid iref in pi.references)
          {
            Guid iff = pi.dependencies.Where(g => g == iref).FirstOrDefault();
            if (iff == Guid.Empty)
            {
              MissingDependency item = new MissingDependency();
              item.Text = pi.fileInfo.Name;
              item.SubItems.Add("Refences to " + projects.Where(p => p.Value.guid == iref).First().Value.fileInfo.Name + " but not depends on it.");
              item.project = pi.guid;
              item.missingDep = iref;
              listView1.Items.Add(item);
            }
          }
          foreach (Guid idep in pi.dependencies)
          {
            if (pi.references.Where(g => g == idep).FirstOrDefault() == Guid.Empty)
            {
              ListViewItem item = new ListViewItem(pi.fileInfo.Name);
              item.SubItems.Add("Depends on " + projects.Where(p => p.Value.guid == idep).First().Value.fileInfo.Name + " but not references to it.");
              listView1.Items.Add(item);
            }
          }
        }
        toolStripStatusLabel1.Text = projects.Count.ToString() + " projects";
        listView1.Refresh();
      }
      catch (Exception ex)
      {
        if (currentLine > 0)
        {
          MessageBox.Show("Error in line " + currentLine.ToString() + "\n" + ex.Message);
        }
        else
        {
          MessageBox.Show(ex.Message);
        }
      }
    }

    static string FindBinName(FileInfo projInfo)
    {
      StreamReader sr = File.OpenText(projInfo.FullName);
      string name = "";
      while (sr.Peek() >= 0)
      {
        string line = sr.ReadLine();
        if (line.Contains("<AssemblyName>"))
        {
          name = line.Substring(line.IndexOf(">") + 1, line.IndexOf("</") - line.IndexOf(">") - 1);
          name = name.Split(new char[] { '\\', '/' }).Last();
          break;
        }
        if (line.Contains("<OutputFile>"))
        {
          if (line.Contains("</OutputFile>"))
          {
            string prjName = projInfo.Name.Substring(0, projInfo.Name.Length - projInfo.Extension.Length);
            name = line.Substring(line.IndexOf(">") + 1, line.IndexOf("</") - line.IndexOf(">") - 1);
            name = name.Split(new char[] { '\\', '/' }).Last();
            name = name.Replace("$(IntDir)", "");
            name = name.Replace("$(OutDir)", "");
            name = name.Replace("$(ProjectName)", prjName);
            break;
          }
          else
          {
            System.Diagnostics.Debug.Assert(false, "</OutputFile>");
          }
        }
      }
      if (name == "")
      {
        name = projInfo.Name;
      }
      name = name.ToLower();
      name = name.Replace(".exe", "");
      name = name.Replace(".dll", "");
      name = name.Replace(".lib", "");
      name = name.Replace(".vcxproj", "");
      sr.Close();
      return name;
    }

    private bool AddDependency(ref List<string> lines, Guid project, Guid missingDep)
    {
      bool inProject = false;
      Guid guid = Guid.Empty;
      for (int i = 0; i < lines.Count; i++)
      {
        string line = lines[i];
        line = line.Trim();
        if (inProject)
        {
          if (line == "ProjectSection(ProjectDependencies) = postProject")
          {
            if (guid == project)
            {
              lines.Insert(i + 1, "\t\t{" + missingDep.ToString() + "} = {" + missingDep.ToString() + "}");
              return true;
            }
          }
          if (line == "EndProject")
          {
            if (guid == project)
            {
              lines.Insert(i, "\tEndProjectSection");
              lines.Insert(i, "\t\t{" + missingDep.ToString() + "} = {" + missingDep.ToString() + "}");
              lines.Insert(i, "\tProjectSection(ProjectDependencies) = postProject");
              return true;
            }
            guid = Guid.Empty;
            inProject = false;
          }
        }
        else if (line.Contains("Project(\"{"))
        {
          inProject = true;
          string[] names = line.Split('=')[1].Split(',');
          string projGuid = names[2];

          projGuid = projGuid.Trim(new char[] { '\\', '"', ' ' });

          guid = new Guid(projGuid);
        }
      }
      return false;
    }

    private void toolStripButton3_Click(object sender, EventArgs e)
    {
      StreamReader sr = File.OpenText(slnFileName);
      List<string> lines = new List<string>();
      while (sr.Peek() >= 0)
      {
        lines.Add(sr.ReadLine());
      }
      sr.Close();
      foreach (var item in listView1.Items)
      {
        MissingDependency problem = item as MissingDependency;
        if (problem != null)
        {
          bool ok = AddDependency(ref lines, problem.project, problem.missingDep);
        }
      }
      StreamWriter sw = File.CreateText(slnFileName + ".new");
      foreach (string line in lines)
      {
        sw.WriteLine(line);
      }
      sw.Close();
    }

    private void toolStripButton4_Click(object sender, EventArgs e)
    {
      LoadSln();
    }

    private void toolStripButton5_Click(object sender, EventArgs e)
    {
      listView1.Items.Clear();
      foreach (ProjectInfo pi in projects.Values)
      {
        try
        {
          string text = File.ReadAllText(pi.fileInfo.FullName, Encoding.UTF8/*Encoding.GetEncoding(1252)*/);
          if (pi.fileInfo.DirectoryName.ToLower().StartsWith("f:"))
          {
            int directiries = pi.fileInfo.DirectoryName.Split(new char[] { '\\', '/' }).Length;
            string dirUp = "";
            for (int i = 1; i < directiries; i++)
            {
              dirUp += "..\\";
            }
            //text = text.Replace("<Import Project=\"f:\\solutions\\RSA.Paths.$(Configuration).$(Platform).props\" />",
            //                    "<Import Project=\"" + dirUp + "solutions\\RSA.Paths.$(Configuration).$(Platform).props\" />");
            //text = text.Replace(@"F:\solutions\RSA.NativeMinimumRules.ruleset",
            //                    dirUp + "solutions\\RSA.NativeMinimumRules.ruleset");
            //if (text.Contains("<ModuleDefinitionFile>") && text.Contains("</ModuleDefinitionFile>"))
            //{
            //  int t1 = text.IndexOf("<ModuleDefinitionFile>");
            //  int t2 = text.IndexOf("</ModuleDefinitionFile>");
            //  if (t2 > t1)
            //  {
            //    string filename = text.Substring(t1 + 22, t2 - t1 - 22);
            //    if (filename.ToLower().StartsWith("f:") && File.Exists(filename))
            //    {
            //      text = text.Replace(filename, filename.Split(new char[] { '/', '\\' }).Last().ToLower());
            //    }
            //  }
            //}

            //text = text.Replace("f:/c", dirUp);
            if (!text.Contains("UpToRootDir"))
            {
              text = text.Replace("<ImportGroup Label=\"ExtensionSettings\">", "<ImportGroup Label=\"ExtensionSettings\"></ImportGroup><PropertyGroup><UpToRootDir>" + dirUp + "</UpToRootDir></PropertyGroup><ImportGroup>");
              File.WriteAllText(pi.fileInfo.FullName, text, Encoding.UTF8/*Encoding.GetEncoding(1252)*/);
            }
          }
        }
        catch (Exception ex)
        {
          listView1.Items.Add("Exception: " + ex.Message);
        }
      }
    }

    private void toolStripButton6_Click(object sender, EventArgs e)
    {
      SaveFileDialog saveExtensFile = new SaveFileDialog();
      saveExtensFile.Filter = "Text Files (*.txt)|*.txt";
      if (saveExtensFile.ShowDialog() == DialogResult.OK)
      {
        StreamWriter sw = File.CreateText(saveExtensFile.FileName);
        sw.WriteLine("graph {");
        Dictionary<Guid, string> guid2bin = new Dictionary<Guid, string>();
        foreach (ProjectInfo project in projects.Values)
        {
          guid2bin[project.guid] = project.binName;
        }
        foreach (ProjectInfo project in projects.Values)
        {
          foreach (Guid guid in project.references)
          {
            sw.WriteLine(project.binName + " -- " + guid2bin[guid] + ";");
          }
        }
        sw.WriteLine("}");
        sw.Close();
      }
    }
  }

  internal class MissingDependency : ListViewItem
  {
    public Guid project;
    public Guid missingDep;
  }

  internal class ProjectInfo
  {
    public Guid guid;
    public string binName;
    public FileInfo fileInfo;
    public List<Guid> dependencies;
    public List<Guid> references;
  }

}
