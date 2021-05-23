using Azure.AI.FormRecognizer.Models;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace luval.ai.docs.msft.analyzer
{
    public interface IInvoiceAnalyzer
    {
        IEnumerable<RecognizedForm> Execute(Stream document);
        Task<IEnumerable<RecognizedForm>> ExecuteAsync(Stream document, CancellationToken cancellationToken);
    }
}