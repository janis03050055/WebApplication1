﻿// <auto-generated />
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Storage;
using Microsoft.EntityFrameworkCore.Storage.Internal;
using System;
using WebApplication1.Models;

namespace WebApplication1.Migrations
{
    [DbContext(typeof(WebApplication1Context))]
    [Migration("20180307051812_Initial")]
    partial class Initial
    {
        protected override void BuildTargetModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "2.0.1-rtm-125")
                .HasAnnotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn);

            modelBuilder.Entity("WebApplication1.Models.Calendar", b =>
                {
                    b.Property<int>("ID")
                        .ValueGeneratedOnAdd();

                    b.Property<DateTime>("date");

                    b.Property<string>("description");

                    b.Property<string>("holidayCategory")
                        .IsRequired();

                    b.Property<string>("isHoliday")
                        .IsRequired();

                    b.Property<string>("name");

                    b.HasKey("ID");

                    b.ToTable("Calendar");
                });
#pragma warning restore 612, 618
        }
    }
}