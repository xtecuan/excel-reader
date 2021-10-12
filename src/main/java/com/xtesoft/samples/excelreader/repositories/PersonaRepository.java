package com.xtesoft.samples.excelreader.repositories;

import com.xtesoft.samples.excelreader.entities.Persona;
import org.springframework.data.jpa.repository.JpaRepository;

public interface PersonaRepository extends JpaRepository<Persona, Long> {
}
