using System;
using System.Collections.Generic;
using System.Text;

namespace JiangLiQuery.IServices
{
    public interface IService<T> where T:class
    {
        IEnumerable<T> Query();

        T Query(int id);

        T Install(T newModel);

        T Modify(T newModel);

        T Delete(int id);

    }
}
