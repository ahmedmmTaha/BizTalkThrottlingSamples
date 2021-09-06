namespace THSample.DFM.PipelineComponents
{
    using System;
    using System.IO;
    using System.Text;
    using System.Drawing;
    using System.Resources;
    using System.Reflection;
    using System.Diagnostics;
    using System.Collections;
    using System.ComponentModel;
    using Microsoft.BizTalk.Message.Interop;
    using Microsoft.BizTalk.Component.Interop;
    using Microsoft.BizTalk.Component;
    using Microsoft.BizTalk.Messaging;
    using Microsoft.BizTalk.Streaming;
    using Microsoft.RuleEngine;
    using System.Xml;
    using System.Runtime.Caching;

    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [System.Runtime.InteropServices.Guid("9373abf9-e608-4ff4-8ba7-95872fc90480")]
    [ComponentCategory(CategoryTypes.CATID_Any)]
    public class ThrottlingPipelineComponent : BaseCustomTypeDescriptor, Microsoft.BizTalk.Component.Interop.IComponent, IBaseComponent, IPersistPropertyBag, IComponentUI
    {
        private static ResourceManager resourceManager;
        private static MemoryCache cache = MemoryCache.Default;
        private bool _Enabled;

        [Browsable(true)]
        public bool Enabled
        {
            get
            {
                return _Enabled;
            }
            set
            {
                _Enabled = value;
            }
        }

        private int _TPM;

        [Browsable(true)]
        [BtsPropertyName("Property_TPM_Name")]
        [BtsDescription("Property_TPM_Description")]
        public int StaticTPM
        {
            get
            {
                return _TPM;
            }
            set
            {
                _TPM = value;
            }
        }

        private string _ThrottleBREPolicyName;

        [Browsable(true)]
        [BtsPropertyName("Property_ThrottleBREPolicyName_Name")]
        [BtsDescription("Property_ThrottleBREPolicyName_Description")]
        public string ThrottleBREPolicyName
        {
            get
            {
                return _ThrottleBREPolicyName;
            }
            set
            {
                _ThrottleBREPolicyName = value;
            }
        }

        private string _ConfigsBREPolicyName;

        [Browsable(true)]
        [BtsPropertyName("Property_ConfigsBREPolicyName_Name")]
        [BtsDescription("Property_ConfigsBREPolicyName_Description")]
        public string ConfigsBREPolicyName
        {
            get
            {
                return _ConfigsBREPolicyName;
            }
            set
            {
                _ConfigsBREPolicyName = value;
            }
        }


        private int _AgencyID;

        [Browsable(true)]
        [BtsPropertyName("Property_AgencyID_Name")]
        [BtsDescription("Property_AgencyID_Description")]
        public int AgencyID
        {
            get
            {
                return _AgencyID;
            }
            set
            {
                _AgencyID = value;
            }
        }

        static ThrottlingPipelineComponent()
        {
            resourceManager = new ResourceManager("THSample.DFM.PipelineComponents.ThrottlingPipelineComponent", Assembly.GetExecutingAssembly());
        }
        public ThrottlingPipelineComponent()
            : base(resourceManager)
        {

        }


        #region IBaseComponent members
        /// <summary>
        /// Name of the component
        /// </summary>
        [Browsable(false)]
        public string Name
        {
            get
            {
                return resourceManager.GetString("COMPONENTNAME", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Version of the component
        /// </summary>
        [Browsable(false)]
        public string Version
        {
            get
            {
                return resourceManager.GetString("COMPONENTVERSION", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Description of the component
        /// </summary>
        [Browsable(false)]
        public string Description
        {
            get
            {
                return resourceManager.GetString("COMPONENTDESCRIPTION", System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        #endregion

        #region IPersistPropertyBag members
        /// <summary>
        /// Gets class ID of component for usage from unmanaged code.
        /// </summary>
        /// <param name="classid">
        /// Class ID of the component
        /// </param>
        public void GetClassID(out System.Guid classid)
        {
            classid = new System.Guid("9373abf9-e608-4ff4-8ba7-95872fc90480");
        }

        /// <summary>
        /// not implemented
        /// </summary>
        public void InitNew()
        {
        }

        /// <summary>
        /// Loads configuration properties for the component
        /// </summary>
        /// <param name="pb">Configuration property bag</param>
        /// <param name="errlog">Error status</param>
        public virtual void Load(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, int errlog)
        {
            object val = null;
            val = this.ReadPropertyBag(pb, "StaticTPM");
            if ((val != null))
            {
                this._TPM = ((int)(val));
            }
            val = this.ReadPropertyBag(pb, "ThrottleBREPolicyName");
            if ((val != null))
            {
                this._ThrottleBREPolicyName = ((string)(val));
            }
            val = this.ReadPropertyBag(pb, "ConfigsBREPolicyName");
            if ((val != null))
            {
                this._ConfigsBREPolicyName = ((string)(val));
            }
            val = this.ReadPropertyBag(pb, "AgencyID");
            if ((val != null))
            {
                this._AgencyID = ((int)(val));
            }
            val = this.ReadPropertyBag(pb, "Enabled");
            if ((val != null))
            {
                this._Enabled = ((bool)(val));
            }
        }

        /// <summary>
        /// Saves the current component configuration into the property bag
        /// </summary>
        /// <param name="pb">Configuration property bag</param>
        /// <param name="fClearDirty">not used</param>
        /// <param name="fSaveAllProperties">not used</param>
        public virtual void Save(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, bool fClearDirty, bool fSaveAllProperties)
        {
            this.WritePropertyBag(pb, "StaticTPM", this.StaticTPM);
            this.WritePropertyBag(pb, "ThrottleBREPolicyName", this.ThrottleBREPolicyName);
            this.WritePropertyBag(pb, "ConfigsBREPolicyName", this.ConfigsBREPolicyName);
            this.WritePropertyBag(pb, "AgencyID", this.AgencyID);
            this.WritePropertyBag(pb, "Enabled", this.Enabled);
        }

        #region utility functionality
        /// <summary>
        /// Reads property value from property bag
        /// </summary>
        /// <param name="pb">Property bag</param>
        /// <param name="propName">Name of property</param>
        /// <returns>Value of the property</returns>
        private object ReadPropertyBag(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, string propName)
        {
            object val = null;
            try
            {
                pb.Read(propName, out val, 0);
            }
            catch (System.ArgumentException)
            {
                return val;
            }
            catch (System.Exception e)
            {
                throw new System.ApplicationException(e.Message);
            }
            return val;
        }

        /// <summary>
        /// Writes property values into a property bag.
        /// </summary>
        /// <param name="pb">Property bag.</param>
        /// <param name="propName">Name of property.</param>
        /// <param name="val">Value of property.</param>
        private void WritePropertyBag(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, string propName, object val)
        {
            try
            {
                pb.Write(propName, ref val);
            }
            catch (System.Exception e)
            {
                throw new System.ApplicationException(e.Message);
            }
        }
        #endregion
        #endregion

        #region IComponentUI members
        /// <summary>
        /// Component icon to use in BizTalk Editor
        /// </summary>
        [Browsable(false)]
        public IntPtr Icon
        {
            get
            {
                return ((System.Drawing.Bitmap)(resourceManager.GetObject("COMPONENTICON", System.Globalization.CultureInfo.InvariantCulture))).GetHicon();
            }
        }

        /// <summary>
        /// The Validate method is called by the BizTalk Editor during the build 
        /// of a BizTalk project.
        /// </summary>
        /// <param name="obj">An Object containing the configuration properties.</param>
        /// <returns>The IEnumerator enables the caller to enumerate through a collection of strings containing error messages. These error messages appear as compiler error messages. To report successful property validation, the method should return an empty enumerator.</returns>
        public System.Collections.IEnumerator Validate(object obj)
        {
            // example implementation:
            // ArrayList errorList = new ArrayList();
            // errorList.Add("This is a compiler error");
            // return errorList.GetEnumerator();
            return null;
        }
        #endregion

        #region IComponent members
        /// <summary>
        /// Implements IComponent.Execute method.
        /// </summary>
        /// <param name="pc">Pipeline context</param>
        /// <param name="inmsg">Input message</param>
        /// <returns>Original input message</returns>
        /// <remarks>
        /// IComponent.Execute method is used to initiate
        /// the processing of the message in this pipeline component.
        /// </remarks>
        public Microsoft.BizTalk.Message.Interop.IBaseMessage Execute(Microsoft.BizTalk.Component.Interop.IPipelineContext pc, Microsoft.BizTalk.Message.Interop.IBaseMessage inmsg)
        {
            try
            {
                Guid callToken = Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceIn();

                long scopeStarted;
                scopeStarted = Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceStartScope("EndToEndScope", callToken);

                string ThrottlePromotedPropertyName = resourceManager.GetString("ThrottlePromotedPropertyName");
                string NINPromotedPropertyName = resourceManager.GetString("NINPromotedPropertyName");
                string AgencyIDPromotedPropertyName = resourceManager.GetString("AgencyIDPromotedPropertyName");
                string PropertySchemaNamespace = resourceManager.GetString("PropertySchemaNamespace");

                string NIN = inmsg.Context.Read(NINPromotedPropertyName, PropertySchemaNamespace).ToString();
                string AgencyRequestThrottleFlag = inmsg.Context.Read(ThrottlePromotedPropertyName, PropertySchemaNamespace).ToString();
                int AID = Convert.ToInt32(inmsg.Context.Read(AgencyIDPromotedPropertyName, PropertySchemaNamespace));

                Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("Throttling Pipeline Enabled Mode: {0}", Enabled);
                Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("Agency Request Throttling Property: {0}", AgencyRequestThrottleFlag);

                if (AgencyID == 0) // If Agency ID is not set (i.e. 0), will use the AgencyID: "AID" promoted property on the message
                    AgencyID = AID;

                if (AgencyRequestThrottleFlag == "1" && this.Enabled)
                {
                    int latencyMilliseconds;
                    int SettingsExpiry_Minutes;
                    int WaitDuringUndefinedRange_Minutes;

                    if (StaticTPM > 0) // Static TPM Will have high priority
                    {
                        Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("Static TPM was set explicitly in Pipeline with value: {0}", StaticTPM);
                        latencyMilliseconds = Convert.ToInt32((60d / _TPM) * 1000);
                    }
                    else // BRE Agency-Configured TPM
                    {
                        Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("Dynamic TPM detected, Will check BRE for Agency-Configured TPM value");
                        string cacheThrottleKey = "DFM_AID_" + AgencyID;
                        object obj = cache.Get(cacheThrottleKey);

                        // Calling DFM BRE Configs
                        scopeStarted = Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceStartScope("CallBRE_Configs", callToken);
                        CallBREConfigsPolicy(out SettingsExpiry_Minutes, out WaitDuringUndefinedRange_Minutes);
                        Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceEndScope("CallBRE_Configs", scopeStarted, callToken);
                        //

                        Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("DFM BRE Configs, SettingsExpiryInMinutes: {0}, WaitDuringUndefinedRangeInMinutes: {1}", SettingsExpiry_Minutes, WaitDuringUndefinedRange_Minutes);
                        Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("Check for Control Message: Will Force New BRE Throttle Call, NIN = : {0}", NIN);

                        if (obj == null)
                        {
                            CacheItemPolicy policy = new CacheItemPolicy();
                            policy.AbsoluteExpiration = DateTime.Now.AddMinutes(SettingsExpiry_Minutes);
                            Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("Cache Expiry in Minutes = {0}", SettingsExpiry_Minutes);

                            // Calling Throttling BRE
                            scopeStarted = Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceStartScope("CallBRE_Throttle", callToken);
                            latencyMilliseconds = CallBREThrottlingPolicy();
                            Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceEndScope("CallBRE_Throttle", scopeStarted, callToken);
                            //

                            Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("BRE Throttle Latency_MilliSeconds = {0}", latencyMilliseconds);

                            cache.Set(cacheThrottleKey, latencyMilliseconds, policy, null);
                        }
                        else
                        {
                            latencyMilliseconds = Convert.ToInt32(obj);
                            Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("Cached Throttle Latency_MilliSeconds = {0}", latencyMilliseconds);
                        }
                    }

                    //Skip if Latency is set in BRE to Unlimited (0) 
                    if (latencyMilliseconds > 0)
                    {
                        Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceInfo("Thread Sleeping now with latency = {0}", latencyMilliseconds);
                        //System.Threading.Thread.Sleep(latencyMilliseconds);
                        Sleep(latencyMilliseconds);
                    }
                    // Retry in Undefined range (i.e not sent by BRE)
                    else if (latencyMilliseconds < 0)
                    {
                        Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceWarning("BRE Throttle Undefined - Detected: Latency_MilliSeconds = {0}", latencyMilliseconds);
                        #region Retry during Undefined Range
                        while (latencyMilliseconds < 0) // Retry during Undefined ranges until the TPM <> -1 then send the first queued Message afterwards immediately
                        {
                            // Calling DFM BRE Configs
                            scopeStarted = Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceStartScope("CallBRE_Configs", callToken);
                            CallBREConfigsPolicy(out SettingsExpiry_Minutes, out WaitDuringUndefinedRange_Minutes);
                            Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceEndScope("CallBRE_Configs", scopeStarted, callToken);
                            //

                            Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceWarning("BRE Throttle Undefined Latency_MilliSeconds = {0}, Will Start waiting for {1} Minutes between retries", latencyMilliseconds, WaitDuringUndefinedRange_Minutes);

                            //Sleep - Wait during undefined Range
                            System.Threading.Thread.Sleep(new TimeSpan(0, WaitDuringUndefinedRange_Minutes, 0));
                            //

                            // Calling Throttling BRE
                            scopeStarted = Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceStartScope("CallBRE_Throttle", callToken);
                            latencyMilliseconds = CallBREThrottlingPolicy();
                            Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceEndScope("CallBRE_Throttle", scopeStarted, callToken);
                            //
                            if (latencyMilliseconds >= 0) // Breaking Condition (it is no more -1: OFF)
                            {
                                string cacheThrottleKey = "DFM_AID_" + AgencyID;
                                CacheItemPolicy policy = new CacheItemPolicy();
                                policy.AbsoluteExpiration = DateTime.Now.AddMinutes(SettingsExpiry_Minutes);
                                cache.Set(cacheThrottleKey, latencyMilliseconds, policy, null);
                            }
                        }
                        #endregion
                    }
                }

                Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceEndScope("EndToEndScope", scopeStarted, callToken);
                Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceOut(callToken);

                // Reset the Throttle Flag on the Request before sending it
                inmsg.Context.Promote(ThrottlePromotedPropertyName, PropertySchemaNamespace, "0");

                return inmsg;
            }
            catch (PolicyExecutionException polEx)
            {
                // Log the inner exception.
                Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceError(polEx);
                THSample.DFM.Utilities.Helper.Log(polEx.InnerException.ToString(), Utilities.MSGType.Error, Utilities.EventSource.THSample_DFM_Pipeline_Components);
                return null;
            }
            catch (RuleEngineException)
            {
                Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceError("Rule Engine exception: Policy could not be instantiated.");
                THSample.DFM.Utilities.Helper.Log("Rule Engine exception: Policy could not be instantiated.", Utilities.MSGType.Error, Utilities.EventSource.THSample_DFM_Pipeline_Components);
                return null;
            }
            catch (Exception ex)
            {
                Microsoft.BizTalk.CAT.BestPractices.Framework.Instrumentation.TraceManager.PipelineComponent.TraceError(ex);
                THSample.DFM.Utilities.Helper.Log(ex.Message, Utilities.MSGType.Error, Utilities.EventSource.THSample_DFM_Pipeline_Components);
                return null;
            }
        }

        private void Sleep(int SleepinMilliseconds)
        {
            DateTime startTime = DateTime.Now;
            while (true)
            {
                Double elapsedMillisecs = ((TimeSpan)(DateTime.Now - startTime)).TotalMilliseconds;
                if (elapsedMillisecs >= SleepinMilliseconds)
                    break;
            }
        }

        private int CallBREThrottlingPolicy()
        {
            int latencyMilliseconds;
            int mTPM;
            DateTime dt = DateTime.Now;

            string str_XMLTemplate = resourceManager.GetString("ThrottlingXML_Str_Template");
            StringBuilder str_XMLDoc = new StringBuilder();
            str_XMLDoc.AppendFormat(str_XMLTemplate, AgencyID, dt.ToString("yyyy-MM-dd"), Convert.ToInt32(dt.TimeOfDay.TotalMinutes), ("*" + dt.DayOfWeek + "*"));

            XmlReader reader = XmlReader.Create(new StringReader(str_XMLDoc.ToString()));
            TypedXmlDocument txdoc = new TypedXmlDocument("THSample.DFM.InternalSchemas.ThrottleSettings", reader);

            System.Data.SqlClient.SqlConnection SQLConn = new System.Data.SqlClient.SqlConnection();
            SQLConn.ConnectionString = THSample.DFM.Utilities.Helper.getTHSampleDBConnectionString();
            SQLConn.Open();

            Microsoft.RuleEngine.DataConnection AgencyThrottleSettingsDataConnection = new Microsoft.RuleEngine.DataConnection("THSample.DFM", "AgencyThrottleSettings", SQLConn);
            Microsoft.RuleEngine.DataConnection AgencyDataConnection = new Microsoft.RuleEngine.DataConnection("THSample.DFM", "Agency", SQLConn);

            object[] facts = new object[3];
            facts[0] = AgencyDataConnection;
            facts[1] = AgencyThrottleSettingsDataConnection;
            facts[2] = txdoc;

            //using-statement to ensure that Dispose method is always called after execution, even if Exception occurs
            using (var throttlingPolicy = new Policy(ThrottleBREPolicyName))
            {
                throttlingPolicy.Execute(facts);
                AgencyDataConnection.Update();
                SQLConn.Close();
            }

            mTPM = txdoc.GetInt32(resourceManager.GetString("ThrottlingXML_TPMNode_XPATH"));
            if (mTPM <= 0) // if TPM is OFF: -1 or Unlimited: 0, send it as-is in Latency so that Throttling can handle them accordingly
            {
                latencyMilliseconds = mTPM;
            }
            else
            {
                latencyMilliseconds = Convert.ToInt32((60.0d / mTPM) * 1000);
                if (latencyMilliseconds <= 70) latencyMilliseconds = Convert.ToInt32( latencyMilliseconds * 0.75);
            }
            return latencyMilliseconds;
        }
        private void CallBREConfigsPolicy(out int SettingsExpiry_Minutes, out int WaitDuringUndefinedRange_Minutes)
        {
            SettingsExpiry_Minutes = 0;
            WaitDuringUndefinedRange_Minutes = 0;

            string str_XMLTemplate = resourceManager.GetString("ConfigsXML_Str_Template");
            StringBuilder str_XMLDoc = new StringBuilder(str_XMLTemplate);

            XmlReader reader = XmlReader.Create(new StringReader(str_XMLDoc.ToString()));
            TypedXmlDocument txdoc = new TypedXmlDocument("THSample.DFM.InternalSchemas.GlobalConfigs", reader);

            //using-statement to ensure that Dispose method is always called after execution, even if Exception occurs
            using (var configsPolicy = new Policy(ConfigsBREPolicyName))
            {
                configsPolicy.Execute(txdoc);
            }

            SettingsExpiry_Minutes = txdoc.GetInt32(resourceManager.GetString("SettingsExpiry_XPATH"));
            WaitDuringUndefinedRange_Minutes = txdoc.GetInt32(resourceManager.GetString("WaitDuringUndefinedRange_XPATH"));
        }
        #endregion
    }
}
