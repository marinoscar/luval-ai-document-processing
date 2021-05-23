﻿using Azure.AI.FormRecognizer;
using Azure.AI.FormRecognizer.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace luval.ai.docs.msft.analyzer
{
    public class InvoiceAnalyzer
    {

        public InvoiceAnalyzer(FormRecognizerClient client) : this(client, new RecognizeInvoicesOptions() { Locale = "en-US" })
        {

        }

        public InvoiceAnalyzer(FormRecognizerClient client, RecognizeInvoicesOptions options)
        {
            Client = client;
            Options = options;
        }

        protected virtual FormRecognizerClient Client { get; private set; }
        protected virtual RecognizeInvoicesOptions Options { get; private set; }

        public async Task<IEnumerable<RecognizedForm>> ExecuteAsync(Stream document, CancellationToken cancellationToken)
        {
            RecognizedFormCollection invoices = await Client.StartRecognizeInvoicesAsync(document, Options, cancellationToken).WaitForCompletionAsync();
            return invoices.ToList();
        }

        public IEnumerable<RecognizedForm> Execute(Stream document)
        {
            var res = ExecuteAsync(document, CancellationToken.None);
            res.Wait();
            return res.Result;
        }
    }
}
