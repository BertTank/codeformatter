// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.ComponentModel.Composition;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Formatting;
using Microsoft.CodeAnalysis.CSharp;
using System.Linq;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Collections.Generic;
using System.Collections.Immutable;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.Options;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.Host;
using System.Collections;
using System.Reflection;

namespace Microsoft.DotNet.CodeFormatting.Rules
{
	[LocalSemanticRule(FormatDocumentFormattingRule.Name, FormatDocumentFormattingRule.Description, LocalSemanticRuleOrder.IsFormattedFormattingRule)]
	internal sealed class FormatDocumentFormattingRule : ILocalSemanticFormattingRule
	{
		internal const string Name = "FormatDocument";
		internal const string Description = "Run the language specific formatter on every document";

		private readonly Options _options;

		private static Type s_IFormattingRuleType = Type.GetType("Microsoft.CodeAnalysis.Formatting.Rules.IFormattingRule, Microsoft.CodeAnalysis.Workspaces");
		private static Type s_ISyntaxFormattingServiceType = Type.GetType("Microsoft.CodeAnalysis.Formatting.ISyntaxFormattingService, Microsoft.CodeAnalysis.Workspaces");
		private static Type s_ListIFormattingRuleType = null;  // type can just be used via Reflection in runtime

		private static MethodInfo s_GetDefaultFormattingRulesMethod = typeof(Formatter).GetMethod("GetDefaultFormattingRules", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic, null, new Type[] { typeof(Workspace), typeof(string) }, null);
		private static MethodInfo s_GetServiceMethod = typeof(HostLanguageServices).GetMethod("GetService");
		private static MethodInfo s_ListAddIFormattingRuleMethod = null; // type can just be used via Reflection in runtime
		private static MethodInfo s_GetFormattedRootMethod = null; // type can just be used via Reflection in runtime
		private static MethodInfo s_FormatMethod = null; // type can just be used via Reflection in runtime

		[ImportingConstructor]
		internal FormatDocumentFormattingRule(Options options)
		{
			_options = options;
		}

		public bool SupportsLanguage(string languageName)
		{
			return
				languageName == LanguageNames.CSharp ||
				languageName == LanguageNames.VisualBasic;
		}

		public async Task<SyntaxNode> ProcessAsync(Document document, SyntaxNode syntaxNode, CancellationToken cancellationToken)
		{
			document = await FormatAsyncEx(document, cancellationToken: cancellationToken);
			// {org} document = await Formatter.FormatAsync(document, cancellationToken: cancellationToken);

			if (!_options.PreprocessorConfigurations.IsDefaultOrEmpty)
			{
				var project = document.Project;
				var parseOptions = document.Project.ParseOptions;
				foreach (var configuration in _options.PreprocessorConfigurations)
				{
					var list = new List<string>(configuration.Length + 1);
					list.AddRange(configuration);
					list.Add(FormattingEngineImplementation.TablePreprocessorSymbolName);

					var newParseOptions = WithPreprocessorSymbols(parseOptions, list);
					document = project.WithParseOptions(newParseOptions).GetDocument(document.Id);
					document = await Formatter.FormatAsync(document, cancellationToken: cancellationToken);
				}
			}

			return await document.GetSyntaxRootAsync(cancellationToken);
		}

		private static ParseOptions WithPreprocessorSymbols(ParseOptions parseOptions, List<string> symbols)
		{
			var csharpParseOptions = parseOptions as CSharpParseOptions;
			if (csharpParseOptions != null)
				return csharpParseOptions.WithPreprocessorSymbols(symbols);

			var basicParseOptions = parseOptions as VisualBasicParseOptions;
			if (basicParseOptions != null)
				return basicParseOptions.WithPreprocessorSymbols(symbols.Select(x => new KeyValuePair<string, object>(x, true)));

			throw new NotSupportedException();
		}

		private static async Task<Document> FormatAsyncEx(Document document, OptionSet options = null, CancellationToken cancellationToken = default(CancellationToken))
		{
			if (document == null)
				throw new ArgumentNullException("document");

			SyntaxNode node = await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(false);
			return document.WithSyntaxRoot(FormatEx(node, document.Project.Solution.Workspace, options, cancellationToken));
		}

		private static SyntaxNode FormatEx(SyntaxNode node, Workspace workspace, OptionSet options = null, CancellationToken cancellationToken = default(CancellationToken))
		{
			return FormatEx(node, workspace, options, null, cancellationToken);
		}

		internal static SyntaxNode FormatEx(SyntaxNode node, Workspace workspace, OptionSet options, IEnumerable<object> rules, CancellationToken cancellationToken)
		{
			if (workspace == null)
				throw new ArgumentNullException("workspace");

			if (node == null)
				throw new ArgumentNullException("node");

			return FormatEx(node, new TextSpan[] { node.FullSpan }, workspace, options, rules, cancellationToken);
		}

		internal static SyntaxNode FormatEx(SyntaxNode node, IEnumerable<TextSpan> spans, Workspace workspace, OptionSet options, IEnumerable<object> rules, CancellationToken cancellationToken)
		{
			if (workspace == null)
				throw new ArgumentNullException("workspace");

			if (node == null)
				throw new ArgumentNullException("node");

			if (spans == null)
				throw new ArgumentNullException("spans");

			object service = getLanguageService(workspace, node);

			if (service != null)
			{
				options = (options ?? workspace.Options);

				// edit behavior: use tabs instead of spaces!
				if (options != null)
					options = options.WithChangedOption(FormattingOptions.UseTabs, "C#", true);

				rules = (rules ?? getDefaultFormattingRules(workspace, node));
				return invokeFormat(node, spans, options, rules, cancellationToken, service);
			}
			return node;
		}

		private static SyntaxNode invokeFormat(SyntaxNode node, IEnumerable<TextSpan> spans, OptionSet options, IEnumerable<object> rules, CancellationToken cancellationToken, object service)
		{
			// return service.Format(node, spans, options, rules, cancellationToken).GetFormattedRoot(cancellationToken);

			if (s_ListIFormattingRuleType == null)
			{
				var listType = typeof(List<>);
				s_ListIFormattingRuleType = listType.MakeGenericType(s_IFormattingRuleType);
			}

			var genericEnumerable = Activator.CreateInstance(s_ListIFormattingRuleType);

			if (s_ListAddIFormattingRuleMethod == null)
				s_ListAddIFormattingRuleMethod = s_ListIFormattingRuleType.GetMethod("Add");

			foreach (var rule in rules)
				s_ListAddIFormattingRuleMethod.Invoke(genericEnumerable, new object[] { rule });

			var args = new object[] { node, spans, options, genericEnumerable, cancellationToken };

			if (s_FormatMethod == null)
				s_FormatMethod = service.GetType().GetMethod("Format");

			var formatResult = s_FormatMethod.Invoke(service, args);

			if (s_GetFormattedRootMethod == null)
				s_GetFormattedRootMethod = formatResult.GetType().GetMethod("GetFormattedRoot");

			args = new object[] { cancellationToken };
			var formattedRootResult = s_GetFormattedRootMethod.Invoke(formatResult, args);

			return (SyntaxNode)formattedRootResult;
		}

		private static IEnumerable<object> getDefaultFormattingRules(Workspace workspace, SyntaxNode node)
		{
			// Formatter.GetDefaultFormattingRules(workspace, node.Language);
			object rules = s_GetDefaultFormattingRulesMethod.Invoke(null, new object[] { workspace, node.Language });
			return (rules as IEnumerable).OfType<object>();
		}

		private static object getLanguageService(Workspace workspace, SyntaxNode node)
		{
			// ISyntaxFormattingService is internal!
			//ILanguageService service = workspace.Services.GetLanguageServices(node.Language).GetService<ILanguageService>();

			var services = workspace.Services.GetLanguageServices(node.Language);

			var gmi = s_GetServiceMethod.MakeGenericMethod(s_ISyntaxFormattingServiceType);

			return gmi.Invoke(services, null);
		}

	}
}