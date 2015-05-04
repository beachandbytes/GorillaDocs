using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs
{
    public class ObjectMapper
    {
        public static void MergeNulls<T>(T Source, T Destination)
        {
            Mapper.CreateMap<T, T>().ForAllMembers(opt => opt.Condition(t => t.DestinationValue == null || t.DestinationValue == ""));
            //Mapper.CreateMap<T, T>().ForAllMembers(opt =>
            //    {
            //        opt.Condition(t => t.DestinationValue == null || t.DestinationValue == "" );
            //    });
            //opt.Condition(t => (t.DestinationType == typeof(string) && t.DestinationValue == ""));
            //(t.DestinationType == typeof(string) && t.DestinationValue == string.Empty)));
            Mapper.Map(Source, Destination);
        }

        public static void Copy<T>(T Source, T Destination)
        {
            Mapper.CreateMap<T, T>();
            Mapper.Map(Source, Destination);
        }

        public static void MergeCollections<T>(IList<T> Source, IList<T> Destination)
        {
            for (int i = 0; i < Source.Count; i++)
                if (i < Destination.Count)
                    ObjectMapper.MergeNulls(Source[i], Destination[i]);
                else
                    Destination.Add(Source[i]);
        }
    }
}

//Action<Person, Person> MapAction = (source, destination) =>
//{
//    if (string.IsNullOrEmpty(destination.FirstName))
//        destination.FirstName = source.FirstName;
//    if (string.IsNullOrEmpty(destination.LastName))
//        destination.LastName = source.LastName;
//};

//Mapper.CreateMap<Person, Person>().ForAllMembers(opt => opt.Condition(srs => !srs.IsSourceValueNull));
//Mapper.CreateMap<Person, Person>().ForAllMembers(opt => opt.UseDestinationValue());
//Mapper.CreateMap<Person, Person>().ForAllMembers(opt =>
//    {
//        opt.UseDestinationValue();
//        opt.Ignore();
//    });

//Mapper.CreateMap<Person, Person>()
//   .AfterMap((s, d) => { MapDetailsAction(s, d); })
//   .ForMember(dest => dest.Details, opt => opt.UseDestinationValue());



//Action<Person, Person> MapDetailsAction = (source, destination) =>
//{
//    if (destination.Details == null)
//    {
//        destination.Details = new Details();
//        destination.Details =
//            Mapper.Map<ItemViewModel, Item>(
//            source.Details, destination.Details);
//    }
//};

//Mapper.CreateMap<Person, Person>()
//    .ForMember(
//            destination => destination.LastName,
//            option =>
//            {
//                option.Condition(rc =>
//                {
//                    var profileViewModel = (Person)rc.InstanceCache.First().Value;
//                    return string.IsNullOrEmpty(profileViewModel.LastName);
//                });

//                option.MapFrom(source => source.LastName);
//            }
//    );

