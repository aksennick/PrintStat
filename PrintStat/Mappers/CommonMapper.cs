using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using AutoMapper;
using PrintStat.Models.ViewModels;

namespace PrintStat.Mappers
{
    public class CommonMapper: IMapper
    {
        static CommonMapper()
        {
            Mapper.CreateMap<DeviceConsumableView, Device>();
            Mapper.CreateMap<Device, DeviceConsumableView>();
            Mapper.CreateMap<ProfileView, Profile>();
         //       .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Name));
            //Mapper.CreateMap<ProfileView, Settings>()
            //    .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.NameParametr));
            //Mapper.CreateMap<ProfileView, SettingValue>()
            //    .ForMember(dest => dest.Value, opt => opt.MapFrom(src => src.Value));

            Mapper.CreateMap< Profile,ProfileView>();
            //    .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Name));
            //Mapper.CreateMap<Settings, ProfileView>()
            //    .ForMember(dest => dest.NameParametr, opt => opt.MapFrom(src => src.Name));
            //Mapper.CreateMap<SettingValue, ProfileView>()
            //    .ForMember(dest => dest.Value, opt => opt.MapFrom(src => src.Value));

                
            Mapper.CreateMap<SizePaperView, SizePaper>()
               .ForMember(dest => dest.Height_cm, opt => opt.MapFrom(src => src.Height_cm))
               .ForMember(dest => dest.Width_cm, opt => opt.MapFrom(src => src.Width_cm));
            Mapper.CreateMap<SizePaper, SizePaperView>()
               .ForMember(dest => dest.Height_cm, opt => opt.MapFrom(src => src.Height_cm))
               .ForMember(dest => dest.Width_cm, opt => opt.MapFrom(src => src.Width_cm));

            Mapper.CreateMap<DeviceType, DeviceTypeView>();
            Mapper.CreateMap<DeviceTypeView, DeviceType>();

            Mapper.CreateMap<PrintKind, PrintKindView>();
            Mapper.CreateMap<PrintKindView, PrintKind>();

            Mapper.CreateMap<Job, JobView>();
            Mapper.CreateMap<JobView, Job>();

            Mapper.CreateMap<Device, DeviceView>()
                .ForMember(dest=>dest.InvNumber,opt=>opt.MapFrom(src=>src.InvNumber))
                .ForMember(dest => dest.ModelID, opt => opt.MapFrom(src => src.ModelID))
                .ForMember(dest => dest.SearchString, opt => opt.MapFrom(src => src.SearchString))
                .ForMember(dest => dest.Sn, opt => opt.MapFrom(src => src.Sn))
                .ForMember(dest => dest.StatisticsSupported, opt => opt.MapFrom(src => src.StatisticsSupported))
                .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Name));
            Mapper.CreateMap<DeviceView, Device>()
                .ForMember(dest=>dest.InvNumber,opt=>opt.MapFrom(src=>src.InvNumber))
                .ForMember(dest => dest.ModelID, opt => opt.MapFrom(src => src.ModelID))
                .ForMember(dest => dest.SearchString, opt => opt.MapFrom(src => src.SearchString))
                .ForMember(dest => dest.Sn, opt => opt.MapFrom(src => src.Sn))
                .ForMember(dest => dest.StatisticsSupported, opt => opt.MapFrom(src => src.StatisticsSupported))
                .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Name));

            Mapper.CreateMap<PaperType, PaperTypeView>();
            Mapper.CreateMap<PaperTypeView,PaperType>();

            Mapper.CreateMap<CartridgeColor, CartridgeColorView>();
            Mapper.CreateMap<CartridgeColorView, CartridgeColor>();

            Mapper.CreateMap<Application, ApplicationView>();
            Mapper.CreateMap<ApplicationView, Application>();

            Mapper.CreateMap<Department, DepartmentView>();
            Mapper.CreateMap<DepartmentView, Department>();

            Mapper.CreateMap<Employee, EmployeeView>();
            Mapper.CreateMap<EmployeeView, Employee>();

            //Mapper.CreateMap<SettingsView, Setup>();
            //Mapper.CreateMap<Setup, SettingsView>();

            Mapper.CreateMap<Settings,FieldSettingView>();
            Mapper.CreateMap<FieldSettingView, Settings>();

            Mapper.CreateMap<Tag, TagView>()
                .ForMember(dest => dest.Tag1, opt => opt.MapFrom(src => src.Tag1))
                .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Name));
            Mapper.CreateMap<TagView, Tag>()
                .ForMember(dest => dest.Tag1, opt => opt.MapFrom(src => src.Tag1))
                .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Name));

            Mapper.CreateMap<TagType, TagTypeView>();
            Mapper.CreateMap<TagTypeView, TagType>();

            Mapper.CreateMap<Component, ComponentView>();
            Mapper.CreateMap<ComponentView,Component>();

            Mapper.CreateMap<TypeConsumable, TypeConsumableView>();
            Mapper.CreateMap<TypeConsumableView, TypeConsumable>();

            Mapper.CreateMap<Consumable, ConsumableView>();
            Mapper.CreateMap<ConsumableView, Consumable>();

            Mapper.CreateMap<Manufacturer, ManufacturerView>();
            Mapper.CreateMap<ManufacturerView, Manufacturer>();

            Mapper.CreateMap<Model, ModelView>();
            Mapper.CreateMap<ModelView, Model>();
            ////////////////////////////////////
            /// modelView 
            /// 
            //Mapper.CreateMap<Tag, ModelView>()
            //    .ForMember(dest => dest.ID, opt => opt.MapFrom(src => src.ID))
            //    .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Name));
            //Mapper.CreateMap<ModelView,Tag>()
            //    .ForMember(dest => dest.ID, opt => opt.MapFrom(src => src.ID))
            //    .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Name));
            ///////////////////////////////////

        }


        public object Map(object source, Type sourceType, Type destinationType)
        {
            return Mapper.Map(source, sourceType, destinationType);
        }

        public object Map(object source, object destination, Type sourceType, Type destinationType)
        {
            return Mapper.Map(source, destination, sourceType, destinationType);
        
        }
    }
}