using Dapper;
using RadugaMassPrint.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RadugaMassPrint
{
    internal class MySqlManager : IDisposable
    {
        private MySqlConnection _connection = new MySqlConnection();

        internal MySqlManager(string connectionString)
        {
            _connection.ConnectionString = connectionString;
        }

        internal async Task<IEnumerable<DocumentData>> GetFilesFolder(IEnumerable<int> docs_id, DateTime dateFrom, IEnumerable<int> operators, int accountType)
        {
            string sqlCommand =
                $"""
                SELECT 
                    d.name as DocumentName,
                    aa.address as Address,
                    ab.name as BuildingName,
                    ac.name as AccountName,
                    a.agrm_id as AgrmID,
                    a.number as AgreementNumber,
                    o.file_name as FileName,
                    o.doc_id as DocID,
                    o.order_date as OrderDate,
                    CAST(o.curr_summ as Decimal(15,2)) as Sum
                FROM orders o
                JOIN agreements a on o.agrm_id = a.agrm_id
                JOIN accounts ac on ac.uid = a.uid
                JOIN accounts_addr aa on aa.uid = ac.uid
                JOIN documents d on d.doc_id = o.doc_id
                LEFT OUTER JOIN address_building ab on aa.building = ab.record_id
                LEFT OUTER JOIN agreements_addons_vals aav 
                    on aav.agrm_id = o.agrm_id
                    and aav.name in ('detailing', 'transcript')
                    and aav.str_value = 'Да'
                WHERE 
                    (
                        o.doc_id in @docs_id
                        { (docs_id.Count() >= 2 ? "or (o.doc_id in (67,68) and aav.agrm_id is not null)" : "") }
                    )
                    AND aa.type = 2
                    AND ifnull(o.file_name, '') != ''
                    AND o.period >= @dateStart
                    AND o.period < @dateEnd
                    AND o.oper_id in @operators
                    AND ac.type = @acType
                    AND NOT 
                        (
                            d.doc_id = 74
                            AND o.oper_id = 1
                        )
                    AND a.state = 0
                """;

            var sqlArg = new
            {
                docs_id = docs_id,
                dateStart = dateFrom,
                dateEnd = new DateTime(dateFrom.Year, dateFrom.Month + 1, 1),
                operators = operators,
                acType = accountType
            };

            var documentsData = await _connection.QueryAsync<DocumentData>(sqlCommand, sqlArg);
            return documentsData;
        }

        public void Dispose()
        {
            _connection.Dispose();
        }
    }
}
